
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCColumssPresent_)
_ODBCColumssPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcHandleCl.inc'),once
  include('dynstr.inc'),once

! -----------------------------------------------------------  
! queue to store the columns used in a query
! -----------------------------------------------------------  
ColumnsQueue queue,type
ColId          long           ! ord of the column in the select statement
ColName        string(eColNameLength)  ! name of the col, not used by the odbc but by the classes
ColType        SQLSMALLINT    ! type of the col, odbc const
ColValue       SQLPOINTER     ! pointer to the buffer for the result
ColSize        SQLLEN         ! size of the column in bytes
             end 

! --------------------------------------------------------------------- 
! class to manage the colums used in a query or a stored procedure call
! each column used is added to the queue and will be mapped or bound 
! to the destination.  the destenation can be any buffer needed/wanted,
! a queue, a group, an individual field, ...
! --------------------------------------------------------------------  
columnsClass class,type,module('odbcColumnsCl.clw'),link('odbcColumnsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
colQ           &ColumnsQueue,private
fieldList      &IDynStr
! set up and clean up functions
construct      procedure()
init           procedure(),byte,proc
kill           procedure()
destruct       procedure()
! general worker 
clearQ         procedure()
! binds the colums to the query
bindColumns    procedure(long  hStmt),sqlReturn,virtual
! not used 
getFields      procedure(),*cstring,virtual
! -----------------------------------------------------------------
! add the various data types to the queue
! -----------------------------------------------------------------
AddColumn      procedure(*long colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *long colPtr),sqlReturn,proc
!
AddColumn      procedure(*string colPtr),sqlReturn,proc     
AddColumn      procedure(string colLabel, *string colPtr),sqlReturn,proc     
!
AddColumn      procedure(*real colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *real colPtr),sqlReturn,proc
!
AddColumn      procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
!
AddColumn      procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength),sqlReturn,private,proc
! not yet used
addField       procedure(string colLabel),protected,virtual
            end 

  _EndOfInclude_
