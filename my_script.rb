# -*- coding: utf-8 -*
#
require 'time'
require 'io/console'
require 'tempfile'
require 'mail'
require 'byebug'

require 'fileutils'
require 'google/apis/calendar_v3'
require 'googleauth'
require 'googleauth/stores/file_token_store'

$logger = Logger.new("./log/invoice.log.#{Time.now.strftime("%Y%m%d")}")
$gmail_pwd = ''

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
SCOPE = Google::Apis::CalendarV3::AUTH_CALENDAR_READONLY
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "invoice_script.yaml")
CLIENT_SECRETS_PATH = 'client_secrets.json'
APPLICATION_NAME = 'Ruby Invoice Script'
GOOGLE_LOGIN = 'nique.rio'
USAGE = "ruby my_script.rb [fourDigitYear-twoDigitMonth]"
TEACHING_CALENDAR_ID='ustsm6tni91g6b9pautfr3vi6c@group.calendar.google.com'

def get_gmail_pwd
  #Get Gmail Password
  puts "Gmail Password for " + GOOGLE_LOGIN
  system "stty -echo" 
  $stdout.flush
  $gmail_pwd = $stdin.gets.chomp
  system "stty echo" 
end

def authorize
  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(
    client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(
      base_url: OOB_URI)
    puts "Open the following URL in the browser and enter the " +
         "resulting code after authorization"
    puts url
    code = $stdin.gets.chomp
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI)
  end
  credentials
end

class Student
  def initialize(name,lesson_type,parent,rate,email)
    @name, @lesson_type, @parent, @email = name, lesson_type, parent, email
    @rate = rate.to_i
    @lessons = Array.new
    @amount_due = 0.0
    @teacher_email = 'mrio@umich.edu'
    @from_address = 'mrio@umich.edu'
    @email_subject = ''
    @email_body = ''
  end

  attr_reader :name, :lesson_type, :parent, :rate, :email, :amount_due, :lessons

  def add_lesson(start,finish)
    @lessons.push( { start: start, finish: finish} )
  end

  def run
    send_email
  end


  private


  def find_amount_due
    sum = 0
    @lessons.each do |lesson| 
      sum += (lesson[:finish] - lesson[:start])/3600
    end
    @amount_due = (sum * @rate)
  end

  def create_message
    #Creates Email Message based on info from calendar and info from csv file.
    #Requires @lessons to be populated.


    #Subject Determined by @lesson_type
    if @lesson_type == 'Piano'
      @email_subject = "Piano Teaching Invoice for #{@lessons[0][:start].strftime("%B")}"
    elsif @lesson_type == 'Tutoring'
      @email_subject = "Tutoring Invoice for #{@lessons[0][:start].strftime("%B")}"
    end

    #Intro for Parents Different than for Adults.
    if @name != @parent
      @email_body += "Hi #{@parent},\n\nHere's what I have on the calendar for #{@name} " \
	"for #{@lessons[0][:start].strftime("%B")}:\n\n"
    else
      @email_body += "Hi #{@parent},\n\nHere's what I have on the calendar for you " \
	"for #{@lessons[0][:start].strftime("%B")}:\n\n"
    end

    #Sort Lessons by Date
    sorted_lessons = @lessons.sort_by{ |lesson| lesson[:start] }
    sorted_lessons.each do |lesson|
      @email_body += lesson[:start].strftime("%A, %B %d from %I:%M - ")  
      @email_body += lesson[:finish].strftime("%I:%M\n") 
    end

    find_amount_due

    @email_body += "\nTotal due for #{@lessons[0][:start].strftime("%B")}: $%.2f.\n\n" \
      "See You Soon-\n-Monique\n" % @amount_due

  end

  def send_email
    flag = true

    create_message

    while flag do
      puts @email
      puts @email_body
      puts "What do you want to do? Send: s; Edit e; Skip: n "
      to_do = STDIN.getch

      if to_do == 's' 
        mail = Mail.new

        mail[:from] = @teacher_email 
        mail[:to] = @email 
        mail[:cc] = @teacher_email 
        mail[:subject] =  @email_subject
        mail[:body]  =   @email_body
        tries = 0;
        begin
          tries += 1
          mail.delivery_method :smtp, { 
            :address              => "smtp.gmail.com",
            :port                 => 587,
            :domain               => 'localhost',
            :user_name            => GOOGLE_LOGIN,
            :password             => $gmail_pwd,
            :authentication       => 'plain',
            :enable_starttls_auto => true  }
          mail.deliver!
        rescue Net::SMTPAuthenticationError => e
          if tries < 3
            puts e.message 
            puts "Wrong Gmail Password. Try again"
            get_gmail_pwd
            retry
          else
            puts e.message 
            abort("Too many tries")
          end
        end

        $logger.info("SENT message to #{@email} <#{@parent}>\n#{@email_body}\n")
        flag = false #Done with this student

      elsif to_do == 'e'
        temp = Tempfile.new('invoice')
        temp.write(@email_body)
        temp.flush
        system("vim #{temp.path}")

        temp.rewind
        @email_body = temp.read
        temp.close
        temp.unlink

        flag = true #try again with updated email_body
      elsif to_do == 'n'
        $logger.info("SKIPPED message for #{@email} <#{@parent}>")
        flag = false #Done with this student

      else
        flag = true #try again since you entered something stupid
      end

    end
  end

end

#-------------------------------------
#              MAIN
#-------------------------------------

#Hash that holds instances of Student; Key: studentNameString; Value: Student;
students = Hash.new

#Process CSV file
IO.foreach("students.txt") do |s| 
  s.strip!
  row = s.split(';')
  students[row[0]] = Student.new(*row) 
end

#Deal with first Argument

first_arg = ARGV[0]

if !first_arg then abort(USAGE) end

#Process Year Month Command Line Arguments

year,month = first_arg.split('-')
year = year.to_i
month = month.to_i

if month > 12 || month < 1 then abort(USAGE) end
if month == 12 
  next_month = 1
  next_year = year+1
else
  next_month = month + 1
  next_year = year
end

start_date = Time.new(year, month)
end_date = Time.new(next_year, next_month)



##Parse other Command Line Arguments; Not really needed for Ruby version

get_gmail_pwd

#Auth with Google
service = Google::Apis::CalendarV3::CalendarService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

params = {
          order_by: 'startTime',
          show_deleted: false,
          single_events: true,
          time_min: start_date.strftime("%FT%T%:z"),
          time_max: end_date.strftime("%FT%T%:z"), 
         }

#Get lesson data from Google Calendar
result = service.list_events(TEACHING_CALENDAR_ID, params)

page_token = nil
while true
  events = result.items
  events.each do |e|
    title = e.summary;
    puts title
    title = title.gsub(/ Lesson/, '') 
    title = title.gsub(/ Tutoring/, '') 
    time_s =  e.start.date_time.to_s
    time_e =  e.end.date_time.to_s
    students[title].add_lesson(Time.parse(time_s),Time.parse(time_e))
  end
  if !(page_token=result.next_page_token)
    break
  end
  result = client.execute(:api_method => calendar.events.list, 
                          :parameters => params.merge({'pageToken' => page_token}),
                         )
end

#Iterate over students. Sends Email 'n' stuff for each instance of students.
students.each do |key,s|
  if !s.lessons.empty?
    s.run
  end
end

$logger.close

