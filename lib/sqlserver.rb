require 'win32ole'

class SqlServer
  attr_accessor :connection, :data, :fields
  attr_writer :username, :password
  def initialize(host, username = 'sa', password='')
    @connection = nil
    @data = nil
    @host = host
    @username = username
    @password = password
  end
  def open(database)
    connection_string =  "Provider=SQLOLEDB.1;"
    connection_string << "Persist Security Info=False;"
    connection_string << "User ID=#{@username};"
    connection_string << "password=#{@password};"
    connection_string << "Initial Catalog=#{database};"
    connection_string << "Data Source=#{@host};"
    connection_string << "Network Library=dbmssocn"
    @connection = WIN32OLE.new('ADODB.Connection')
    @connection.Open(connection_string)
  end
  def query(sql)
    recordset = WIN32OLE.new('ADODB.Recordset')
    recordset.Open(sql, @connection)
    @fields = []
    recordset.Fields.each do |field|
      @fields << field.Name
    end
    begin
      recordset.MoveFirst
      @data = recordset.GetRows
    rescue
      @data = []
    end
    recordset.Close
    @data = @data.transpose
  end
  def execute_command(sql_statement)
    command = WIN32OLE.new('ADODB.Command')
    command.ActiveConnection = @connection
    command.CommandText=sql_statement
    command.CommandTimeout = 120
    command.Execute
  end
  def close
    @connection.Close
  end
end