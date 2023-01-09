     member('OdbcDemo_Ms.clw')

     map
       module('win32')
         WaitForSingleObject(long hHandle, long dwMilliseconds),bool,pascal,name('WaitForSingleObject')
         CreateEvent(long lpEventAttributes, BOOL bManualReset, BOOL bInitialState, *cstring lpName),long,pascal,raw,name('CreateEventA')
       end
       !callAsyncSp(long connStr)
       GetEventHandle(),long
     end

callAsyncSp procedure(string connString)

query   &ODBCClType
conn    &ODBCConnectionClType
connStr &MSConnStrClType

param   &ParametersClass

SQL_OV_ODBC3_80 long(380)
ev      long,auto
addr long

  code

  addr = connString
  conn &= (addr)

  !conn.Init(connStr)

  query &= new ODBCClType()
  query.Init(conn)
  conn.setOdbcVersion(SQL_OV_ODBC3_80)

  ev = GetEventHandle()
  conn.Connect(ev, true)

  query.execSp('dbo.waitDemo', param)

  WaitForSingleObject(ev, 2000)

  query.endAsync()

  conn.disconnect()

  Message('End it')

  return
! ----------------------------------------------------------------

GetEventHandle procedure()

evname cstring('thisEvent')
retH   long,auto
nullStr &cstring

  code

  retH = CreateEvent(0, false, false, evname)


  return retH
! ---------------------------------------------------------------------------