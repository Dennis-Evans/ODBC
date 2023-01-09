
  member()
  
  include('odbcErrorCl.inc'),once 

  map 
    module('odbc32')
      SQLGetDiagField(SQLSMALLINT HandleType, SQLHANDLE Handle, SQLSMALLINT RecNumber, SQLSMALLINT DiagIdentifier, *SQLPOINTER DiagInfoPtr, SQLSMALLINT BufferLength, *SQLSMALLINT StringLengthPtr),sqlReturn,pascal,name('SQLGetDiagFieldW')
      SQLGetDiagRec(SQLSMALLINT HandleType, SQLHANDLE Handle, SQLSMALLINT RecNumber, odbcWideStr errState, *SQLINTEGER NativeErrorPtr, odbcWideStr MessageText, SQLSMALLINT BufferLength, *SQLSMALLINT TextLengthPtr),sqlReturn,pascal,raw,name('SQLGetDiagRecW')
    end 
  end

! ---------------------------------------------------------------------------

OdbcErrorClType.init procedure() !,sqlReturn

  code 

  if (self.makeObjects() <> sql_Success)
    return sql_error
  end 
  
  return sql_Success
! end init 
! ----------------------------------------------------------------------------
  
OdbcErrorClType.kill procedure() !,virtual

  code 
  
  self.destroyObjects()
  
  return
! end kill
! ----------------------------------------------------------------------------
  
OdbcErrorClType.destruct procedure()

  code 
  
  self.kill()
  
  return
! end destruct 
! ----------------------------------------------------------------------------  
  
OdbcErrorClType.getConnError procedure(ODBCConnectionClType conn) !,sqlReturn,proc

retv   sqlReturn,auto

  code 

  !retv = self.getError(SQL_HANDLE_ENV, conn.gethEnv())
  
  return retv 
! end getConnError
! ----------------------------------------------------------------------------  

OdbcErrorClType.setDisplayError procedure(bool onOff)

  code 
  
  self.displayError = onOff
  
  return
! end setDisplayError 
! ----------------------------------------------------------------------------  
   
OdbcErrorClType.getEnvError procedure(ODBCConnectionClType conn) !,sqlReturn,proc

retv   sqlReturn,auto

  code 

  retv = self.getError(SQL_HANDLE_ENV, conn.gethEnv())
  
  return retv 
! end getDatabaseError
! ----------------------------------------------------------------------------

OdbcErrorClType.getDataBaseError procedure(ODBCConnectionClType conn) !,sqlReturn,proc

retv   sqlReturn,auto

  code 

  retv = self.getError(SQL_HANDLE_DBC, conn.gethDbc())
  
  return retv 
! end getDatabaseError
! ----------------------------------------------------------------------------  
     
OdbcErrorClType.getError procedure(SQLSMALLINT HandleType, SQLHANDLE Handle)  

retv      sqlReturn,auto
errCount  long,auto
count     long,auto

claStateMsg  cstring(12)
claErrMsg    cstring(2001)

stateMsg  CWideStr
errMsg    CWideStr

outState  CStr
outErr    CStr

tempholder bool

  code 
  
  self.freeErrorMsgQ()
  self.getDiagRecCount(handleType, handle)

  loop count = 1 to self.errorCount
    
    claStateMsg = all(' ')
    tempholder = statemsg.Init(claStateMsg)

    claErrMsg = all(' ')
    tempholder = errMsg.Init(claErrMsg)

    retv = SQLGetDiagRec(handleType, handle, count, stateMsg.getWideStr(), self.errorMsgQ.NativeErrorPtr, errMsg.getWideStr(), 2000, self.errorMsgQ.textLengthPtr)
    if (retv = sql_Success) or (retv = sql_success_with_info)
      tempholder = outState.Init(stateMsg)
      self.errorMsgQ.sqlState = outState.getCStr()
      tempholder = outErr.Init(errMsg)
      self.errorMsgQ.MessageText = outErr.getCStr()
      add(self.errorMsgQ)
    end  ! if
  end ! loop

  self.displayError = true
  
  if (retv = sql_Success) and (self.displayError = true)
    self.showError()  
  end 

  return retv
! end getError
! ----------------------------------------------------------------------
  
OdbcErrorClType.showError procedure()

count     long,auto

  code 

  loop count = 1 to self.errorCount
    get(self.errorMsgQ, count)
    message('ODBC Error State  ->' & self.errorMsgQ.sqlState & '|Error Message  ->' & clip(self.errorMsgQ.MessageText), 'Database Error', icon:exclamation)
  end  
  
  return   
! end showError
! ----------------------------------------------------------------------
  
OdbcErrorClType.freeErrorMsgQ procedure() !,private

  code 

  self.errorCount = 0  
  free(self.errorMsgQ)
  clear(self.errorMsgQ)
  
  return   
! end freeErrorMsgQ
! ----------------------------------------------------------------------  
  
OdbcErrorClType.getDiagRecCount procedure(SQLSMALLINT HandleType, SQLHANDLE Handle) !,long,private  

retv            sqlReturn,auto
StringLengthPtr short,auto
 
  code 

  retv = SQLGetDiagField(HandleType, Handle, 0, SQL_DIAG_NUMBER, self.errorCount, 0, StringLengthPtr)
  
  case retv 
  !of SQL_SUCCESS  
    ! valid result, nothing to do
  !of SQL_SUCCESS_WITH_INFO 
    ! not currently using this value, if you need to read strings from the record
    ! use this to verify the buffer lengths 
  of SQL_INVALID_HANDLE
    message('Call to getDiagField with an incorrect handle.', 'Invalid Handle', icon:exclamation)
    self.freeErrorMsgQ()
  of SQL_ERROR 
    message('Call to getDiagField returned SQL_ERROR, verify the handle type.', 'SQL error', icon:exclamation)
    self.freeErrorMsgQ()
  end 
    
  return retv
! end getDiagRecCount
! ----------------------------------------------------------------------  
  
OdbcErrorClType.makeObjects procedure() !,sqlReturn,private

  code 

  self.errorMsgQ &= new(OdbcErrorQueue)
  if (self.errorMsgQ &= null) 
    return sql_Error
  end 
    
  return sql_Success
! end makeObjects 
! -------------------------------------------------------------------------
  
OdbcErrorClType.destroyObjects procedure() !,sqlReturn,private

  code 
  
  if (~self.errorMsgQ &= null) 
    self.freeErrorMsgQ()
    dispose(self.errorMsgQ) 
    self.errorMsgQ &= null
  end 
    
  return
! end destroyObjects 
! -------------------------------------------------------------------------  
