! types and equates 

withStatement    equate(true)
withoutStatement equate(false)

odbcWideStr       equate(long)

dbVendor itemize,pre()
MS         equate
PostGre    equate
Oracle     equate
lastVendor equate    ! place holder 
        end 

SQL_NO_HANDLE     equate(0)
        
TIMESTAMP_STRUCT group,type
year     short
month    short
day      short
hour     short
minute   short
second   short
fraction long 
    end  

SQL_HANDLE_ENV   equate(1)
SQL_HANDLE_DBC   equate(2)
SQL_HANDLE_STMT  equate(3)
SQL_HANDLE_DESC  equate(4)

SQL_HANDLE_TYPE  equate(long)

SQL_DRIVER_NOPROMPT equate(0)
SQL_DRIVER_PROMPT equate(2)

SQL_CLOSE        equate(0)
SQL_UNBIND       equate(2)
SQL_RESET_PARAMS equate(3)

SQL_NO_DATA      equate(100)

SQL_NULL_HANDLE  equate(0)

SQL_NTS          equate(-3)

! 3.5 and better and obviously get only
SQL_ATTR_CONNECTION_DEAD equate(1209)
SQL_CD_TRUE              equate(1)  ! Connection is closed/dead */
SQL_CD_FALSE             equate(0)  ! Connection is open/available */


Sql_Success      equate(0)
Sql_Success_With_Info equate(1)
Sql_Error        equate(-1)
sql_Invalid_handle equate(-2)
sql_Allocate_handle equate(-3)

SQL_STILL_EXECUTING equate(2)

SQL_DIAG_NUMBER  equate(2)

SQLRETURN        equate(short) 
SQLSMALLINT      equate(short)
SqlChar          equate(cstring)
SQLUSMALLINT     equate(ushort)
SqlLen           equate(long)
SQLULEN          equate(ulong)  

SQLHANDLE        equate(long)
SQLHENV          equate(SQLHANDLE)
SQLHDBC          equate(SQLHANDLE) 
SQLHSTMT         equate(SQLHANDLE) 
SQLHWND          equate(SQLHANDLE)

SQLINTEGER       equate(long)
SQLPOINTER       equate(long)

SQL_ATTR_ODBC_VERSION equate(200)
SQL_OV_ODBC3     equate(3)

SQL_C_CHAr       equate(1)

! table type parameter, Server 2008 or better
SQL_SS_TABLE        equate(-153)
SQL_SOPT_SS_BASE    equate(1225)
SQL_SOPT_SS_PARAM_FOCUS equate(SQL_SOPT_SS_BASE+11)
SQL_C_DEFAULT    equate(99)

SQL_CHAR         equate(1)
SQL_INTEGER      equate(4)
SQL_FLOAT        equate(6)
SQL_TYPE_TIMESTAMP equate(93)

SQL_SIGNED_OFFSET equate(-20)
SQL_NO_TOTAL      equate(-4)


SQL_C_SLONG      equate(SQL_INTEGER + SQL_SIGNED_OFFSET)
SQL_C_FLOAT      equate(7)
SQL_C_DOUBLE     equate(8)
SQL_C_TYPE_TIMESTAMP equate(11)

SQL_PARAM_INPUT        equate(1)
SQL_PARAM_OUTPUT       equate(4)
SQL_PARAM_INPUT_OUTPUT equate(2)

SQL_BIND_TYPE           equate(5)
SQL_ATTR_ROW_BIND_TYPE  equate(SQL_BIND_TYPE)
SQL_ATTR_ROW_ARRAY_SIZE equate(27)
SQL_ATTR_ROW_STATUS_PTR equate(25)
SQL_ATTR_ROWS_FETCHED_PTR equate(26)

SQL_DIRECTION     equate(long)
SQL_FETCH_NEXT    equate(1)
SQL_FETCH_FIRST   equate(2)
SQL_FETCH_LAST    equate(3)
SQL_FETCH_PRIOR   equate(4)
SQL_ATTR_CURSOR_SCROLLABLE equate(-1)
SQL_SCROLLABLE    equate(1)

SQL_IS_INTEGER  equate(-6)
SQL_IS_POINTER  equate(-4)

SQL_ATTR_ASYNC_ENABLE        equate(4)
SQL_ASYNC_ENABLE_ON          equate(1)
SQL_ASYNC_ENABLE_OFF         equate(0)
SQL_ATTR_ASYNC_STMT_EVENT    equate(29)

SQL_ATTR_ASYNC_DBC_FUNCTIONS_ENABLE EQUATE(117)
SQL_ASYNC_DBC_ENABLE_ON      EQUATE(1)
   
RETCODE    equate(short)

eSizeReal  equate(8)
eSizeLong  equate(4)

eColNameLength    equate(50)
eParamNameLength  equate(50)

eParamIdChar      equate('@')
eSpaceChar        equate(' ')
eOpenParen        equate('(')
eCloseParaen      equate(')')
eOdbcParamIdChar  equate('?')

eSelectLabel      equate('select ')
eFieldListLabel   equate('@FieldList')

eCallLabel        equate('{{call ')
eScalarCallLabel  equate('{{? = call ')
eMarkerComma      equate('?, ')
eFinalMarker      equate('?)}')