
!ABCIncludeFIle

  OMIT('_EndOfInclude_',_ODBCParamPresent_)
_ODBCParamPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  include('odbcSqlStrCl.inc'),once
    
! ----------------------------------------------------------------------    
! queue to store information for each parameter used in a sp call
! or in a query that accepts parameters
! add each parameter to the queue using the various functions
! the code will bind the parameters before the actual call to the database
! ----------------------------------------------------------------------
ParametersQueue queue,type
paramId           long     ! ordinal of the paremter
paramName         string(eParamNameLength)  ! name of the parameter, not yet used      
InOutType         SQLSMALLINT   ! type of parameter. in or out 
valueType         SQLSMALLINT   ! C type of the data in the parameter
paramType         SQLSMALLINT   ! ODBC type of the parameter 
paraSize          SQLULEN       ! len or size of the parameter
DecimalDigits     SQLSMALLINT   ! number of decimal digits, if needed
ParamPtr          SQLPOINTER    ! pointer to the data in the code
paramLength       SQLLEN        ! length
tableName         cstring(128)  ! name of the table type, in the server, if this is a table valued parameter
                end ! queue 
                
!
ParametersClass class,type,module('odbcParamsCl.clw'),link('odbcParamsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
!bindByName            bool,private
setupFailed           bool,private
paramQ                &ParametersQueue
construct             procedure()
init                  procedure(),byte,proc
kill                  procedure()
destruct              procedure()
clearQ                procedure()
!
bindParameters        procedure(SQLHSTMT hStmt, *long numberRows),sqlReturn
bindParameters        procedure(SQLHSTMT hStmt),sqlReturn
focusTableParameter   procedure(SQLHSTMT hStmt, long ordinal),sqlReturn,proc  
unfocusTableParameter procedure(SQLHSTMT hStmt),sqlReturn,proc

HasParameters         procedure(),bool
! ----------------------------------------------------------
! add the various input parameters
! the overloads with the name input are not yet used but cause no harm if called
! the missing data types will be added
! ----------------------------------------------------------
AddInParameter        procedure(*long colPtr),sqlReturn,proc
AddInParameter        procedure(string pName, *long varPtr),sqlReturn,proc
AddInParameter        procedure(*cstring colPtr),sqlReturn,proc     
AddInParameter        procedure(string pName, *cstring colPtr),sqlReturn,proc     
AddInParameter        procedure(*real colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *real colPtr),sqlReturn,proc
AddInParameter        procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddInParameter        procedure(string pname, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
! add input array parameters
AddlongArray          procedure(long array),sqlReturn,proc
AddRealArray          procedure(long array),sqlReturn,proc
AddCStringArray       procedure(long array, long elementSize),sqlReturn,proc
! add a table valued parameter 
AddTableParameter     procedure(long numberRows, *cstring tableName),sqlReturn,proc
! called internall but can be called by the using code, if needed/wanted
addParameter          procedure(SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc
addParameter          procedure(string paramName, SQLSMALLINT InOutType, SQLSMALLINT ValueType, SQLSMALLINT ParameterType, SQLULEN ColumnSize, SQLSMALLINT DecimalDigits, SQLPOINTER varPtr, SQLLEN BufferLength),byte,proc
! ----------------------------------------------------------------
! add the output types 
! ----------------------------------------------------------------
AddOutParameter       procedure(*long varPtr),sqlReturn,proc
AddOutParameter       procedure(*cstring varPtr),sqlReturn,proc
! these are for the name values, not yet used
FillPlaceHolders      procedure(sqlStrClType sqlCode, long startPos = 1),long
findByParamName       procedure(string pName),sqlReturn
                    end

  _EndOfInclude_
