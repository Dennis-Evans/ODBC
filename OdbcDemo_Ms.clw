    program

   include('odbcConn.inc'),once
   include('abwindow.inc'),once
   include('odbcConnStrCl.inc'),once
   include('odbcColumnsCl.inc'),once
   include('odbcParamsCl.inc'),once
   include('odbcSqlStrCl.inc'),once
   include('odbcCl.inc'),once
   include('odbcTypes.inc'),once
   include('dynStr.inc'),once
   
    map
      main()
      module('callAsync.clw')
        callAsyncSp(string connStr)
      end
    end

DemoQuery  class(odbcClType)
formatRow    procedure(),virtual
fillResult   procedure(*columnsClass cols, *queue q),sqlReturn,protected
execTableSp  procedure(string spName, *ParametersClass param, long numberRows),sqlReturn,virtual
fillInsertQueue procedure(),long
           end

databaseGroup group,type
sysId           long
label           string(60)
amount          real
              end 

dataaseGroupOne group,type
label             string(60)
amount             real
                 end 
dbQueue queue(databaseGroup)
        end

dbQueue2 queue(dataaseGroupOne)
        end

insertQueue queue
sysId           long
label           cstring(60)
amount          real
rowAction     long
            end 
  code

  main()

  return

main procedure()

thisWindow class(windowManager)
serverName   string(40)
dbName       string(30)
query        &DemoQuery    
!
msConnStr    &MSConnStrClType
msConn       &OdbcConnectionClType
!
init         procedure(),byte,proc,virtual
takeAccepted procedure(),byte,proc,virtual
ReadTable    procedure()
ReadTableSp  procedure()
ReadTableSpPa procedure(string lbl)
ReadOutSp    procedure()
callScalar   procedure()
readMulti    procedure()
insertTvp    procedure()
callAsync    procedure()
           end

filterLabel    string(60)  

Window WINDOW('Connect'),AT(,,295,278),FONT('MS Sans Serif',8,,FONT:regular),CENTER,GRAY
       PROMPT('Server name:'),AT(33,19),USE(?Prompt2)
       ENTRY(@s30),AT(89,19,117,10),USE(thisWindow.serverName)
       PROMPT('Datbase Name:'),AT(32,39),USE(?Prompt1)
       ENTRY(@s40),AT(88,39,117,10),USE(thisWindow.dbName)
       BUTTON('Table Sp'),AT(226,57,44,14),USE(?btnTableSp)
       BUTTON('Call Async'),AT(17,59,49,14),USE(?btnAsyn)
       BUTTON('C&onnect'),AT(17,78,44,14),USE(?ConnectBtn),DEFAULT
       BUTTON('Read'),AT(93,78,44,14),USE(?btnRead)
       BUTTON('Call Scalar'),AT(161,77,44,14),USE(?btnScalar)
       BUTTON('Multi Results'),AT(221,78,51,14),USE(?btnMulti)
       BUTTON('&Cancel'),AT(229,13,44,14),USE(?CancelBtn)
       BUTTON('Read Sp'),AT(19,99,44,14),USE(?btnReadSp)
       BUTTON('Read Sp Param'),AT(93,101,63,14),USE(?btnReadSpParam)
       BUTTON('Out Sp'),AT(165,100,44,14),USE(?btnOutSp)
       PROMPT('Prompt 3'),AT(223,102,43,10),USE(?Prompt3)
       PROMPT('Label Filter :'),AT(19,128),USE(?Prompt4)
       ENTRY(@s20),AT(63,128,117,10),USE(filterLabel)
       LIST,AT(21,145,254,62),USE(?List1),FORMAT('55L(2)|M~SysId~116L(2)|M~Label~40L(2)|M~Amount~')
       LIST,AT(21,216,254,53),USE(?List2),FORMAT('61L(2)|M~Amount~50L(2)|M~Amount~')
     END

  code

  thisWindow.run()

  return 
! end main 
! ------------------------------------------------------------------------------------

thisWindow.init procedure()

SQL_OV_ODBC3_80 long(380)
retv            byte,auto

  code

  retv = parent.init()

  thisWindow.dbName = 'default_test'
  thisWindow.serverName = 'DENNISHYPERV\DEV'

  self.open(window)

  self.addItem(?CancelBtn, requestCancelled)

  self.msConn &= new(OdbcConnectionClType)
  self.msConnStr &= new (MSConnStrClType)
  self.msConnStr.Init(self.servername, self.dbName)
!  
  if (self.msConn.init(self.msConnStr) = sql_Success) 
    self.msConn.setOdbcVersion(SQL_OV_ODBC3_80)
    self.query &= new DemoQuery()
    self.query.Init(self.msConn)
  end 
    
  return retv
! end init
! -----------------------------------------------------------------------------

thisWindow.takeAccepted procedure()

retv    byte,auto

  code

  retv = parent.takeAccepted()

  case accepted()
  of ?connectBtn
    update()
    retv = SELF.msConn.Connect(withoutStatement)
    if (retv = sql_Success)
      message('Congratulations, a connection to the MS-SQL Server Database has been established.', 'Success', icon:exclamation)
    end
    SELF.msConn.Disconnect(withoutStatement)
    
  of ?btnRead
    self.readTable()
  of ?btnReadSp
    self.ReadTableSp()   
  of ?btnReadSpParam
    self.ReadTableSpPa(filterLabel)
  of ?btnOutSp
    self.readOutSp()
  of ?btnScalar
    self.callScalar()
  of ?btnMulti 
    self.readMulti()
  of ?btnTableSp
    self.insertTvp()
  of ?btnAsyn
    self.callAsync()
  end ! case

  return level:benign
! end takeAccepted
! -----------------------------------------------------------------------------

thisWindow.ReadTable procedure()

cols       columnsClass
retv       sqlReturn
sqlCode    &IDynstr

  code

  free(dbQueue)

  sqlCode &= newdynstr()
  sqlCode.cat('select lb.sysId, lb.label, lb.amount from dbo.labelDemo lb order by lb.Label;')

  cols.AddColumn(dbQueue.sysId)
  cols.AddColumn(dbQueue.label)
  cols.AddColumn(dbQueue.amount)

  if (self.msconn.Connect() = sql_Success) 
    retv = self.query.execQuery(sqlCode, cols, dbQueue)
    self.msConn.Disconnect()
    if (retv = sql_Success)
      ?List1{prop:from} = dbQueue
    end
  end 

  DisposeDynStr(sqlCode)

  return
! ---------------------------------------------------------------------------

demoQuery.formatRow procedure() 

  code

  dbQueue.label = lower(dbQueue.label)
  dbQueue.label[1] = upper(dbQueue.label[1])

  return
! --------------------------------------------------------------------------  

thisWindow.ReadTableSp  procedure()
 
cols          columnsClass
retv          sqlReturn

  code
   
  free(dbQueue) 
  cols.AddColumn(dbQueue.sysId)
  cols.AddColumn(dbQueue.label)
  cols.AddColumn(dbQueue.amount)

  self.msconn.Connect()
  
  retv = self.query.execSp('dbo.ReadLabelDemo', cols, dbQueue)
  
  ?List1{prop:from} = dbQueue

  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------

thisWindow.ReadTableSpPa  procedure(string lbl)
 
cols          columnsClass
retv          sqlReturn
parameters    ParametersClass
param         cstring(200)

  code
   
  free(dbQueue) 

  param = lbl
  cols.AddColumn(dbQueue.sysId)
  cols.AddColumn(dbQueue.label)

  parameters.init()
  retv = parameters.AddInParameter('inLabel', param)
  
  self.msconn.Connect()
  
  retv = self.query.execSp('dbo.ReadLabelDemoByLabel', cols, parameters, dbQueue)
  
  ?List1{prop:from} = dbQueue

  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------  

thisWindow.ReadOutSp procedure()

sqlStr        sqlStrClType  
retv          sqlReturn
parameters    ParametersClass
param         long,auto 
nameparam     cstring(20)

  code

  ?Prompt3{prop:text} = ''
  parameters.init()
  retv = parameters.AddOutParameter(param)

  self.msconn.Connect()
  
  retv = self.query.execSp('dbo.CountDemoLabels', parameters)

  ?Prompt3{prop:text} = param

  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------    

thisWindow.callScalar procedure()

retv          sqlReturn
parameters    ParametersClass
spRetv        long,auto
nameFilter    cstring('Barney')

  code

  ?Prompt3{prop:text} = ''
  parameters.init()
  retv = parameters.AddOutParameter(spRetv)
  retv = parameters.AddInParameter(nameFilter)

  self.msconn.Connect()
  
  retv = self.query.callScalar('dbo.GetId', parameters)

  ?Prompt3{prop:text} = spRetv

  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------    

thisWindow.readMulti procedure()

cols          columnsClass
cols1         columnsClass
retv          sqlReturn

  code
   
  free(dbQueue) 
  free(dbQueue2)

  cols.AddColumn(dbQueue.sysId)
  cols.AddColumn(dbQueue.label)

  cols1.AddColumn(dbQueue2.Amount)
  cols1.AddColumn(dbQueue2.label)

  self.msconn.Connect()
  
  retv = self.query.execSp('dbo.ReadTwo', cols, dbQueue)
  ?List1{prop:from} = dbQueue

  if (retv = sql_Success) 
    if (self.query.nextResultSet() = true) 
      self.query.unbindColums()
      retv = self.query.fillResult(cols1, dbQueue2)
      ?List2{prop:from} = dbQueue2
    end
  end 

  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------  

thisWindow.insertTvp  procedure()

sqlStr        sqlStrClType  
retv          sqlReturn
parameters    ParametersClass
param         long,auto 
tableTypeName cstring('LabelDemoType')
rows          long,auto

  code

  ! add some demo data
  rows = self.query.fillInsertQueue()
  
  ! add the table parameter
  parameters.init()
  retv = parameters.AddTableParameter(rows, tableTypeName)
 
  self.msconn.Connect()  
  retv = self.query.execTableSp('dbo.InsertTable', parameters, rows)
  SELF.msConn.Disconnect()

  return
! --------------------------------------------------------------------------    

! --------------------------------------------------------------------------    
! some notes on this call 
! 
! the number of rows input must remain in scope while the call is being processed.  
! the driver will write back to this value and if it goes out of scope the 
! driver will still write the information back but no one know where it will
! go.  
! the arrays need to reamina in scope.  
! the arrays are declared here so they can be created with the input value.
!
! clarion really needs dynamic arrays and a pointer type but those are 
! never going to happen.  
! --------------------------------------------------------------------------    
DemoQuery.execTableSp procedure(string spName, *ParametersClass param, long numberRows) !,sqlReturn,virtual

retv  sqlReturn

sysIdArray     long,dim(numberRows)
LabelArray     cstring(60),dim(numberRows)
AmountArray    real,dim(numberRows)
RowActionArray long,dim(numberRows)

x              long,auto
tabelvalues    ParametersClass
typeName       cstring('dbo.LabelDemoType')

  code

  if (numberRows <= 0) 
    return sql_error
  end 
  ! fill the arrays from the queue
  loop x = 1 to numberRows
    get(insertQueue, x)
    sysIdArray[x] = insertQueue.sysId
    LabelArray[x] = insertQueue.label  
    AmountArray[x] = insertQueue.Amount
    RowActionArray[x] = insertQueue.rowAction
  end 
  
  retv = self.setupSpCall('dbo.InsertaTable', param)

  retv = param.bindParameters(self.conn.gethStmt(),numberRows)
  
  ! init the parameters for the arrays and set the focus to the table
  tabelvalues.Init()

  tabelvalues.focusTableParameter(self.conn.gethStmt(), 1)
  ! add the arrays and bind 
  tabelvalues.AddlongArray(address(sysIdArray))  
  tabelvalues.AddCStringArray(address(labelArray), size(labelArray[1]))
  tabelvalues.addrealArray(address(amountArray))
  tabelvalues.AddlongArray(address(rowActionArray))  
 
  retv = tabelvalues.bindParameters(self.conn.gethStmt())

  ! remove the focus  and execute
  tabelvalues.unfocusTableParameter(self.conn.gethStmt())

  retv = self.execSp( )

  return retv
! ------------------------------------------------------------------------------------

DemoQuery.fillInsertQueue procedure() !long

retv   long,auto

    code

    insertQueue.SysId = -1
    insertQueue.label = 'Bam Bam'
    insertQueue.Amount = 35.45
    insertQueue.rowAction = 1
    add(insertQueue)

    insertQueue.SysId = -2
    insertQueue.label = 'Deno'
    insertQueue.Amount = 21.33
    insertQueue.rowAction = 1
    add(insertQueue)

!   uncomment the following to update a row and delete a row from the table 
!   in one call

!    insertQueue.SysId = 3
!    insertQueue.label = 'Tom'
!    insertQueue.Amount = 99.99
!    insertQueue.rowAction = 2
!    add(insertQueue)

!    insertQueue.SysId = 2
!    insertQueue.rowAction = 3
!    add(insertQueue)
    
    retv = records(insertQueue)

    return retv
! ------------------------------------------------------------------------------------    

DemoQuery.fillResult   procedure(*columnsClass cols, *queue q) !,sqlReturn,protected

retv   sqlReturn

  code

  retv = parent.fillResult(cols, q)

  return  retv

thisWindow.callAsync    procedure()

addr string(50)

  code

  addr = address(self.msConn)
  start(callAsyncSp, 0, addr)

  return