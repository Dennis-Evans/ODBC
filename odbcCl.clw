

  member()
  
  include('odbcCl.inc'),once
  include('odbcSqlStrCl.inc'),once

  map 
    module('odbc32')
      SQLExecDirect(SQLHSTMT StatementHandle, odbcWideStr StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLExecDirectW')
      SQLExecute(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLCancel(SQLHSTMT StatementHandle),sqlReturn,pascal,proc
      SQLFetch(SQLHSTMT hs),sqlReturn,pascal
      SQLFetchScroll(SQLHSTMT StatementHandle, SQLSMALLINT FetchOrientation, SQLLEN FetchOffset),sqlReturn,pascal
      SQLPrepare(SQLHSTMT StatementHandle, *SQLCHAR StatementText, SQLINTEGER TextLength),sqlReturn,pascal,raw,name('SQLPrepareW')
      SQLCloseCursor(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLMoreResults(SQLHSTMT StatementHandle),sqlReturn,pascal
      SQLFreeStmt(SQLHSTMT StatementHandle, SQLUSMALLINT Option),sqlReturn,pascal
      !SQLSetStmtAttr(SQLHSTMT StatementHandle, SQLINTEGER Attribute, SQLPOINTER ValuePtr, SQLINTEGER StringLength),sqlReturn,pascal
    end
    module('odbcdr')
      SQLCompleteAsync(SQLSMALLINT HandleType, SQLHSTMT Handle,  *RETCODE AsyncRetCodePtr),sqlreturn,pascal,name('SQLCompleteAsync')
    end
  end

! ---------------------------------------------------------------------------
! Init 
! sets up the instance for use.  assigns the connection object input to the 
! data member.  allocates and init's the class to handle the sql statment or string. 
! ---------------------------------------------------------------------------  
odbcClType.init procedure(ODBCConnectionClType conn)   

retv     byte(level:benign)

  code 
  
  if (conn &= null) 
    return level:notify
  end
    
  self.conn &= conn
  self.sqlStr &= new(sqlStrClType)
  if (self.sqlStr &= null)
    retv = level:notify
  else 
    self.sqlStr.init()
  end  
     
  return retv 
! end Init
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! frees the memory used
! ----------------------------------------------------------------------
odbcClType.kill procedure() !,virtual  

  code 

  if (~self.sqlStr &= null)
    self.sqlStr.kill() 
    self.sqlStr &= null
  end  
  
  self.conn &= null
  
  return
! end kill
! ----------------------------------------------------------------------
 
odbcClType.destruct procedure()  

  code 

  self.kill()
  
  return
! end destruct
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! unbinds any columns that are currently bound
! typical use is when a call that returns multiple result sets is in use.
! the first result set is processed and then call this to unbind the columns,
! then bind the caloums for the second result set, 
! repeat as needed
! ----------------------------------------------------------------------
odbcClType.unBindColums procedure()

retv sqlReturn

  code

  ! if therer is a statment handle then unbind
  ! if not then nothing to do
  if (self.conn.getHStmt() > 0) 
    retv = SQLFreeStmt(self.conn.getHStmt(), SQL_UNBIND)
  end

  return
! end unBindColums
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! virtual place holder
! ----------------------------------------------------------------------
odbcClType.formatRow procedure() !,virtual  

  code
  
  ! format queue elements for display in the derived object 
  
  return
! end formatRow 
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! Local worker function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcClType.setSqlCommand procedure(*IdynStr s) ! sqlReturn,protected

  code 
  
  ! make sure there is one
  if (s &= null) or (s.strLen() = 0)
    return sql_Error
  end 
  
  self.sqlStr.init(s)

  return sql_Success
! end setSqlCommand
! ----------------------------------------------------------------------  

! -----------------------------------------------------------------------------
! Local worker overloaded function to assign the sql str (the actual sql statement) used in this 
! call to the class member
! -----------------------------------------------------------------------------  
odbcClType.setSqlCommand procedure(string s) ! sqlReturn,protected

  code 
  
  ! make sure there is one
  if (len(clip(s)) = 0)
    return sql_Error
  end 
  
  self.sqlStr.init(s)

  return sql_Success
! end setSqlCommand
! ----------------------------------------------------------------------  

! ----------------------------------------------------------------------  
! checks for the next result set, if any, and moves to the next result set
! returns true if there is more and false if not
! ----------------------------------------------------------------------  
odbcClType.nextResultSet procedure() 

retv bool(false)
res  sqlReturn

  code 
 
  res = SQLMoreResults(self.conn.getHStmt()) 
  if (res = sql_success) 
    retv = true
  end 
   
  return retv;
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! ------------------------------------------------------------------------------
odbcClType.fetch procedure() !sqlReturn,virtual

retv   sqlReturn

  code 
  
  retv = SQLFetch(self.conn.gethStmt())
  
  if (retv = sql_success_with_info) 
    retv = sql_success
  end 
   
  return retv
! end fetch
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! fetch
! reads the result set, one row at a time and places the data into the queue fields.
! Queue fields are already bound to the columns so all that is needed here is an add(q)
!
! Note, the queue fields must be bound before this method is called.
! ------------------------------------------------------------------------------
odbcClType.fetch procedure(*queue q) !sqlReturn,virtual

retv   sqlReturn
hStmt  SQLHSTMT

  code 
  
  ! start loop and keep looping until an error or no_data is returned  
  ! use a local for the hadle so the function is not called for each row
  hStmt = self.conn.gethStmt()
  
  loop
    retv = SQLFetch(hStmt)
    case retv 
    of SQL_NO_DATA
      ! set back to success, no_data is expected (end of result set), 
      ! but caller is going to check for success
      retv = Sql_Success    
      break
    of Sql_Success
    orof Sql_Success_with_info
      ! format the queue elements for display, if needed, and add the element to the queue
      self.formatRow()
      add(q)
    else 
      ! dump the queue, something went wrong and 
      ! the code should not return a partial result set
      free(q)
      break    
    end  ! case
  end ! loop
  
  if (retv = sql_success_with_info) 
    retv = sql_success
  end 
   
  return retv
! end fetch
! -----------------------------------------------------------------------------
 
! -----------------------------------------------------------------------------
! Binds the columns from the queue to the columns in the result set
! then calls fetch to read the result set
! -----------------------------------------------------------------------------
odbcClType.fillResult procedure(*columnsClass cols, *queue q, long setId = 1) !,sqlReturn,private

retv   sqlReturn 

  code 
 
  ! bind the columns just before the fetch, not needed for the execute query calls 
  ! so do it here, 
  retv = cols.bindColumns(self.conn.getHstmt())

  ! if ok then go fetch the result
  if (retv = sql_success)
    retv = self.fetch(q)
  end  
  
  if (retv <> sql_Success)
    self.getError()
  end   

  return retv
! end fillResult
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
! call the error class to read the error information
! -----------------------------------------------------------------------------
odbcClType.getError procedure() 

retv   sqlReturn
err    ODBCErrorClType

  code 
  
  retv = err.Init()
  err.getError(SQL_HANDLE_STMT, self.conn.getHstmt())

  return
! end getError  
! -----------------------------------------------------------------------------
    
odbcClType.endAsync procedure() !sqlreturn

retv     sqlReturn
outCode  retcode

  code

  retv = SQLCompleteAsync(SQL_HANDLE_STMT, self.conn.getHStmt(), outCode)

  return retv
! -----------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execQuery
! execute a query that returns a result set.  
! get a connection, prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method does not accept the parameters class instance so use this one for queries that 
! do not have parameters.
! ------------------------------------------------------------------------------    
odbcClType.execQuery procedure(*IDynStr sqlCode, *columnsClass cols, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode, cols) <> sql_Success)
    return sql_error
  end 
  
  retv = self.execQuery() 

  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
 
  return retv 
! end execQuery
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execQuery, same as the one above with a different type for the sql code input
! execute a query that returns a result set.  
! get a connection, prep the statement, execute the statement
! then fill the queue or buffers and close the connection when done
!
! this method does not accept the parameters class instance so use this one for queries that 
! do not have parameters.
! ------------------------------------------------------------------------------    
odbcClType.execQuery procedure(*sqlStrClType sqlCode, *columnsClass cols, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode.sqlStr, cols) <> sql_Success)
    return sql_error
  end 
  
   
  retv = self.execQuery() 

  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
 
  return retv 
! end execQuery
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that does not return a result set and does not use any 
! parameters
! ----------------------------------------------------------------------
odbcClType.execQuery procedure(*IDynStr sqlCode) !,sqlReturn,virtual

res     long,auto   ! used t oavoid function call warnings
retv    sqlReturn,auto
wideStr CWideStr

  code 
  
  res = wideStr.Init(sqlCode.Cstr())
  retv = SQLExecDirect(self.conn.gethStmt(), wideStr.GetWideStr(), SQL_NTS)

  return retv
! --------------------------------------------------------------------

! ----------------------------------------------------------------------
! execute a query that returns a result set and expects parameters
! ----------------------------------------------------------------------
odbcClType.execQuery procedure(*IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setupQuery(sqlCode, cols) <> sql_Success)
    return sql_error
  end 

  retv = params.bindParameters(self.conn.gethStmt())

  retv = self.execQuery() 
  
  ! fill the queue
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
  
  return retv 
! end execQuery
! ----------------------------------------------------------------------

odbcClType.execQueryOut procedure(*IDynStr sqlCode, *ParametersClass params) !,sqlReturn,virtual

retv    sqlReturn,auto

  code 
  
  if (self.setSqlCommand(sqlCode) <> sql_Success) 
    return sql_error
  end 
  retv = params.bindParameters(self.conn.gethStmt())

  retv = self.execQuery() 
  
  return retv 
! end execQuery
! ----------------------------------------------------------------------


! ----------------------------------------------------------------------
! execute a query that does not return a result set
! ----------------------------------------------------------------------
odbcClType.execQuery procedure() !,sqlReturn,private

res     long,auto
retv    sqlReturn,auto
wideStr CWideStr

  code 
  
  res = wideStr.Init(self.sqlStr.cstr())
  if (res <= 0) 
    retv = sql_error
  else   
    retv = SQLExecDirect(self.conn.gethStmt(), wideStr.getWideStr(), SQL_NTS)
  
    if (retV = sql_Success_with_info) 
      retv = sql_success
    end 
    if (retv <> sql_success)
      self.getError()
    end
  end 

  return retv
! end execQuery  
! -----------------------------------------------------------------------------
  
! -----------------------------------------------------------------------------
! bind the parameters and then call the execQuery/0 method
! -----------------------------------------------------------------------------  
odbcClType.execQuery procedure(*ParametersClass params) !,sqlReturn,private

retv    sqlReturn(sql_Success)

  code 
  
  ! if none then get out 
  if (params &= null)   
    return sql_error
  end   
    
  retv = params.bindParameters(self.conn.getHStmt())
  
  if (retv = sql_Success)   
    retv = self.execQuery()
  end 
    
  return retv
! -----------------------------------------------------------------------------

! -----------------------------------------------------------------------------
!  virtual place holder
! -----------------------------------------------------------------------------
odbcClType.execTableSp procedure(string spName, *ParametersClass params, long numberRows) !,sqlReturn,virtual

  code
 
  ! this function must be overloaded in a derived class

  return sql_success
! -----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! execute a stored procedure, 
! function calls exec direct with the command text
! -----------------------------------------------------------------------------
odbcClType.execSp procedure() !private,sqlReturn

retv     sqlReturn,auto
wideStr  CWideStr
retCount long  ! used to avoid function warning

  code

  retCount = wideStr.Init(self.sqlStr.cstr())
  retv = SQLExecDirect(self.conn.gethStmt(), wideStr.GetWideStr(), SQL_NTS)
  
  if (retv <> Sql_Success) and (retv <> Sql_Success_with_info)
    self.getError()  
  end
  
  if (retv = Sql_Success_with_info)
    retv = sql_Success  
  end 
    
  return retv
! end execSp
! ---------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName) !,sqlReturn

params  &ParametersClass
retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
  
  if (~params &= null) 
    retv = params.bindParameters(self.conn.gethStmt())
  end
  
  if (retv = sql_Success) 
    retv = self.execSp()
  end  

  return retv 
! end execSp
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that does not return a value or a result set. 
! binds any parameters and calls execSp/0
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, *ParametersClass params) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
  
  if (~params &= null) 
    retv = params.bindParameters(self.conn.gethStmt())
  end
  
  if (retv = sql_Success) 
    retv = self.execSp()
  end  

  return retv 
! end execSp
! ----------------------------------------------------------------------

! ------------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! sp does not expect any parameters
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, columnsClass cols, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName) <> sql_Success)
    return sql_Error
  end 
  
  retv = self.execSp()
  if (retv = sql_Success)
    retv = self.fillResult(cols, q)
  end   
  
  return retv
! end execSp
! ----------------------------------------------------------------------

! -----------------------------------------------------------------------------
! execSp
! call an stored procedure that returns a result set, the 
! queue parameter is bound to the resutls, 
! binds any parameters and calls execSp/0 
! ------------------------------------------------------------------------------  
odbcClType.execSp procedure(string spName, columnsClass cols, *ParametersClass params, *queue q) !,sqlReturn

retv    sqlReturn 

  code 
  
  if (self.setupSpCall(spName, params) <> sql_Success)
    return sql_Error
  end 
    
  retv = params.bindParameters(self.conn.gethStmt())
    
  if (retv = sql_Success) 
    retv = self.execSp() 
    if (retv = sql_Success) 
      retv = self.fillResult(cols, q)
    end   
  end  

  return retv
! end execSp
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! calls a sclar function and puts the returned value in the bound parameter
! ----------------------------------------------------------------------  
odbcClType.callScalar procedure(string spName, *ParametersClass params) 

retv    sqlReturn

  code 
  
  self.sqlStr.formatScalarCall(spName, params)
    
  retv = params.bindParameters(self.conn.gethStmt())
    
  if (retv = sql_Success) 
    retv = self.execSp() 
  end  

  return retv
! end callScalar
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname()}
! this one is used for a stored procedure with no parameters
! ----------------------------------------------------------------------
odbcClType.setupSpCall procedure(string spName) 

retv     sqlReturn,auto
params   &ParametersClass

  code 
  
  retv = self.setupSpCall(spName, params)
  
  return retv 
! end setupSpCall
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! sets up a call, just formats the string with the {call spname(?, ...)}
! this one adds a pllace holder for each parameter
! ----------------------------------------------------------------------  
odbcClType.setupSpCall procedure(string spName, *ParametersClass params) ! sqlReturn,private

retv    sqlReturn 

  code 
  
  if (spName = '') 
    return sql_error
  end 

  if (params &= null) 
    self.sqlStr.formatSpCall(spName)
  else   
    self.sqlStr.formatSpCall(spName, params)
  end   
    
  return retv
! end setupSpCall
! ----------------------------------------------------------------------  

odbcClType.setupQuery procedure(*IDynStr sqlCode, *columnsClass cols) !,sqlReturn,private

  code 
  
  if (self.setSqlCommand(sqlCode) <> sql_Success) 
    return sql_error
  end 
  
  self.sqlStr.replaceFieldList(cols)
 
  return sql_Success  
! end setupQuery
! ----------------------------------------------------------------------  