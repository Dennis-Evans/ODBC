
  member()
  
  include('odbcConnStrCl.inc'),once 

  map 
  end

eTrustedConnTextOn   equate('Trusted_Connection=yes;')
eTrustedConnTextOff  equate('Trusted_Connection=no;')

eDriverName equate('ODBC Driver 13 for SQL Server')

eDriverLabel         equate('Driver={{')
eServerLabel         equate('Server=')
eDbLabel             equate('Database=')

eConnDelimit      equate(';')

! -------------------------------------------------------------------------------------
! Init 
! set up the instance based on the db input and uses the three string input to prime 
! the fields.  
! -------------------------------------------------------------------------------------
baseConnStrClType.init procedure(string srvName, string dbName) !,*IConnString

  code 
  
  self.connStr &= newDynStr()

  self.setDriverName(eDriverName)
  self.setSrvName(srvName)
  self.setDbName(dbName)

  return
! end init 
! -----------------------------------------------------------------
  
! -----------------------------------------------------------------
! kill
! dispose the dyn string 
! -----------------------------------------------------------------
baseConnStrClType.kill procedure()

  code 
  
  disposeDynStr(self.connStr)
  self.connStr &= null
  
  return
! end kill
! -----------------------------------------------------------------

baseConnStrClType.ConnectionString procedure() !,*cstring,virtual

  code 
    
  return null
! end ConnectionString
! ------------------------------------------------------------------------------

! -----------------------------------------------------------------
! Setters for the instance
! -----------------------------------------------------------------
baseConnStrClType.setDriverName procedure(string driverName)

  code 
  
  self.driverName = eDriverLabel & clip(driverName) & '}'  & eConnDelimit
  
  return
! end setDbName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setDbName procedure(string dbname)

  code 
  
  self.dbName = eDbLabel & clip(dbName) & eConnDelimit
  
  return
! end setDriverName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setSrvName procedure(string srvName)

  code 
  
  self.srvName = eServerLabel & clip(srvName) & eConnDelimit
  
  return
! end setServerName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setUserName procedure(string  user)

  code 
  
  self.userName = 'User ID=' & clip(user) & eConnDelimit
  
  return
! end setUserName
! ------------------------------------------------------------------------------
  
baseConnStrClType.setPassword procedure(string pw)

  code 
  
  self.password = 'Password=' & clip(pw) & eConnDelimit
  
  return
! end setpassword
! ------------------------------------------------------------------------------
  
baseConnStrClType.setPortNumber procedure(string portNumber)

  code
  
  self.portNumber = 'Port=' & clip(portNumber) & eConnDelimit
  
  return
! end setPortNumber
! ------------------------------------------------------------------------------

MSConnStrClType.init procedure(string srvName, string dbName) !,virtual

  code

  parent.Init(srvName, dbName)
  self.setTrustedConn(true)

  return
! ------------------------------------------------------------------------------
  
MSConnStrClType.setTrustedConn procedure(bool onOff)

  code 

  self.trustedConn = onOff  
  
  return
! end setTrustedConn
! ------------------------------------------------------------------------------

MSConnStrClType.setUseMars procedure(bool onOff)

  code 
  
  self.useMars = onOff
  
  return 
! end setUseMars
! ------------------------------------------------------------------------------

MSConnStrClType.ConnectionString procedure() !,*cstring,virtual

  code 
  
  ! clear it 
  self.connStr.Kill()
  ! and then build it
  
  ! add the defaults that are always used
  self.connStr.cat(self.driverName & self.SrvName &  self.dbName)

  if (self.TrustedConn = true)
    self.connStr.cat(eTrustedConnTextOn & eConnDelimit)
  else 
    self.connStr.cat(eTrustedConnTextOff)
  end  
     
  return self.connStr.cstr()
! end ConnectionString
! ------------------------------------------------------------------------------