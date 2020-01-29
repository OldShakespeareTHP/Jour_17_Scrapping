# require 'open-uri'
# require 'google_drive'
# require 'csv'

class Scrapper
  attr_accessor :townhall_emails

  def initialize
    @townhall_emails = get_townhall_emails
  end

  def get_email_on_page(townhall_url)
    html = open(townhall_url)
    doc = Nokogiri::HTML(html)
    townhall_emails = doc.css('tbody')[0]
                        .css('tr')[3]
                        .css('td')[1].text
    return townhall_emails
  end

  def get_townhall_pages_urls
    html = open('https://www.annuaire-des-mairies.com/val-d-oise.html')
    doc = Nokogiri::HTML(html)
    townhall_list = doc.css("a[class='lientxt']")
  end

  def get_townhall_emails
    townhall_list = get_townhall_pages_urls
    hash_townhall_emails = Hash.new
    townhall_list.each do |townhall|
      temp = 'http://annuaire-des-mairies.com' + townhall['href'].delete_prefix('.')
      temp2 = adapt_syntax(townhall)
      hash_townhall_emails[temp2] = get_email_on_page(temp)
    end
    return hash_townhall_emails
  end

  def adapt_syntax(word)
    word.text.split.each do |word|
      if (word.size >= 3 && word != 'SUR') || word == 'WY' || word == 'US' || word == 'DIT' || word == 'SOUS'
        word.capitalize!
      else
        word.downcase!
      end
    end.join(' ')
  end

  def save_as_JSON
    File.open("db/emails.json", "w") do |f|
      f.write(JSON.pretty_generate(@townhall_emails))
    end
  end

  def save_as_spreadsheet
    session = GoogleDrive::Session.from_config("config.json")
    worksheet = session.spreadsheet_by_key("1_S9qLvCcMtSDHJJ4XHTKiWO-7Pb8UyP0lABAMn0ewFY").worksheets[0]
    i = 1
    j = 1
    @townhall_emails.each do |k,v|
      worksheet[i, j] = k
      j += 1
      worksheet[i, j] = v
      j -= 1
      i += 1
    end
    worksheet.save  
  end

  def save_as_csv
    CSV.open("db/emails.csv", "w") do |csv|
      @townhall_emails.to_a.each do |elem|
        csv << elem
      end
    end
  end

  def perform
    save_as_JSON
    save_as_spreadsheet
    save_as_csv
  end
end
