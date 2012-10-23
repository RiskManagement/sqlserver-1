SQL Server
=========

Simple gem for accessing MS SQL. This is not new code, you can google and find this. I just wanted to make sure that this was available via rubygems to make development easier.

Usage:

    require 'sqlserver'
    begin
      db = SqlServer.new('server','user','password')
      db.open('Database')
      db.query('SELECT* FROM table')
      db.data.each do |row|
        puts row[0].to_s
      end
      db.close
    rescue Exception => ex
      puts ex.message
      puts ex.backtrace.inspect
    end