require 'sinatra/base'
require 'oauth2'
require 'nokogiri'
require 'microsoft_graph'
require 'time'
require 'tzinfo'
require 'dotenv/load'

ID = ENV['MICROSOFT_GRAPH_ID']
SECRET = ENV['MICROSOFT_GRAPH_SECRET']
SCOPES = ['openid', 'profile', 'User.Read', 'Calendars.ReadWrite']

OAuthSessions = {}

class AppointmentBookings < Sinatra::Base
  enable :sessions

  def oauth_sessions
    OAuthSessions
  end

  def oauth_client
    OAuth2::Client.new(ID, SECRET, {
      site: 'https://login.microsoftonline.com',
      authorize_url: '/common/oauth2/v2.0/authorize',
      token_url: '/common/oauth2/v2.0/token'
    })
  end

  def get_login_url
    oauth_client.auth_code.authorize_url({
      redirect_uri: 'http://localhost:9292/authorize',
      scope: SCOPES.join(' ')
    })
  end

  def get_access_token
    if token_hash = oauth_sessions[session['session_id']]
      token = OAuth2::AccessToken.from_hash(oauth_client, token_hash)

      if token.expired?
        token.refresh!.tap do |token|
          oauth_sessions[session['session_id']] = token.to_hash
        end
      else
        token.token
      end
    else
      false
    end
  end

  get '/' do
    erb :index, locals: { login_url: get_login_url }
  end

  get '/authorize' do
    token = oauth_client.auth_code.get_token(params['code'], {
      redirect_uri: 'http://localhost:9292/authorize',
      scope: SCOPES.join(' ')
    })

    oauth_sessions[session['session_id']] = token.to_hash

    redirect '/book-appointment'
  end

  def graph(token)
    MicrosoftGraph.new({
      base_url: 'https://graph.microsoft.com/v1.0',
      cached_metadata_file: File.join(MicrosoftGraph::CACHED_METADATA_DIRECTORY, 'metadata_v1.0.xml'),
    }, &(->(r) { r.headers['Authorization'] = "Bearer #{token}" }))
  end

  def meetings_calendar(token)
    graph(token).me.calendars.filter(name: 'Meetings').first
  end

  get '/book-appointment' do
    if token = get_access_token
      if calendar = meetings_calendar(token)
        erb :book_appointment, locals: { calendar: calendar }
      else
        'You need to create a "Meetings" calendar'
      end
    else
      redirect '/'
    end
  end

  post '/book-appointment' do
    if token = get_access_token
      if calendar = meetings_calendar(token)
        start_time = Time.parse("#{params['start_date']} #{params['start_time']}")
        end_time = (Time.parse("#{params['start_date']} #{params['start_time']}") + 30 * 60)
        puts start_time.zone

        puts calendar.events.create!({
          Subject: 'Appointment',
          Body: {
            ContentType: 'HTML',
            Body: ''
          },
          Start: {
            DateTime: start_time.iso8601,
            TimeZone: TZInfo.period_for(start_time).identifier
          },
          End: {
            DateTime: end_time.iso8601,
            TimeZone: TZInfo.period_for(end_time).identifier
          },
          Attendees: [{
            EmailAddress: { Address: params['email'] },
            Type: 'Required'
          },]
        }).inspect

        redirect '/book-appointment'
      else
        'You need to create a "Meetings" calendar'
      end
    else
      redirect '/'
    end
  end
end

run AppointmentBookings
