

!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCClassPresent_)
_ODBCClassPresent_   EQUATE(1)

  include('dynStr.inc'),once
  include('odbcConn.inc'),once
  include('odbcTypes.inc'),once
  include('odbcParamsCl.inc'),once 
  include('odbcSqlStrCl.inc'),once
  include('odbcColumnsCl.inc'),once 


odbcClType class,type,module('odbcCl.clw'),link('odbcCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
sqlStr           &sqlStrClType
conn             &ODBCConnectionClType,protected
! set up and tear down
init             procedure(*ODBCConnectionClType conn),byte,proc
kill             procedure(),virtual
destruct         procedure()
unBindColums     procedure()  
! public wirtuals 
formatRow        procedure(),virtual
! public protected
setSqlCommand    procedure(*IdynStr s),sqlReturn,protected
setSqlCommand    procedure(string s),sqlReturn,protected
!
nextResultSet    procedure(),bool
fetch            procedure(),sqlReturn,virtual
fetch            procedure(*queue q),sqlReturn,protected,virtual
fillResult       procedure(*columnsClass cols, *queue q, long setId = 1),sqlReturn,protected
getError         procedure(),private
!
endAsync         procedure(),sqlReturn
! public methods that are used to run a query 
execQuery       procedure(*IDynStr sqlCode),sqlReturn,virtual
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual
execQuery       procedure(*sqlStrClType sqlCode, *columnsClass cols, *queue q),sqlReturn,virtual
execQuery       procedure(*IDynStr sqlCode, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual
execQueryOut    procedure(*IDynStr sqlCode, *ParametersClass params),sqlReturn,virtual
! private query 
execQuery       procedure(),sqlReturn,private
execQuery       procedure(*ParametersClass params),sqlReturn,private
setupQuery      procedure(*IDynStr sqlCode, *columnsClass cols),sqlReturn,private
!
! code to call stored procedures 
execTableSp    procedure(string spName, *ParametersClass params, long numberRows),sqlReturn,virtual
execSp         procedure(string spName),sqlReturn,virtual
execSp         procedure(string spName, *ParametersClass params),sqlReturn,virtual
execSp         procedure(string spName, *columnsClass cols, *queue q),sqlReturn,virtual
execSp         procedure(string spName, *columnsClass cols, *ParametersClass params, *queue q),sqlReturn,virtual
! call a sclar function
callScalar     procedure(string spName, *ParametersClass params),sqlReturn 
! private
execSp         procedure(),protected,sqlReturn
setupSpCall   procedure(string spName),sqlReturn,protected
setupSpCall   procedure(string spName, *ParametersClass params),sqlReturn,protected
            end

  _EndOfInclude_
