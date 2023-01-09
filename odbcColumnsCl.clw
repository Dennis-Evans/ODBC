
  member()
  
  include('odbcColumnsCl.inc'),once 

  map 
    module('odbc32')
      SQLBindCol(SQLHSTMT StatementHandle, SQLUSMALLINT ColumnNumber, SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength, *SQLLEN StrLen_or_Ind),sqlReturn,pascal
    end
  end
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
!  default constructor, calls the init function for the set up
! ---------------------------------------------------------------------------
columnsClass.construct procedure()  

  code 
  
  self.init()
  
  return 
! end construct
! ------------------------------------------------------------------------------
  
! ---------------------------------------------------------------------------
!  allocates the queue and the dyn str used
! ---------------------------------------------------------------------------  
columnsClass.init procedure()

retv      byte(level:benign)

  code 

  ! fields list is not yet used
  self.fieldList &= newDynStr()
  if (self.fieldList &= null) 
    return level:notify
  end 
    
  self.colq &= new(columnsQueue)
  if (self.colQ &= null)
    return level:notify
  end 
    
  return retv
! end init 
! ------------------------------------------------------------------------------
  
! ------------------------------------------------------------------------------
! disposes the queue and the dyn str
! ------------------------------------------------------------------------------
columnsClass.kill procedure()

  code 

  if (~self.colQ &= null)
    free(self.colQ)
    dispose(self.colQ)
    self.colQ &= null
  end    
    
  if (~self.fieldList &= null)
    disposedynstr(self.fieldList)
    self.fieldList &= null
  end
  
  return
! end kill
! ------------------------------------------------------------------------------

! ---------------------------------------------------------------------------
!  default destructor, calls the kill function for the clean up
! ---------------------------------------------------------------------------
columnsClass.destruct procedure()

  code 

  self.kill()  
  
  return 
! end destruct
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! gets the string stored in the dyn str, 
! returns a cstring 
! ------------------------------------------------------------------------------
columnsClass.getFields procedure() !,*cstring,virtual

  code 
  
  return self.fieldList.cStr()
  ! end getFields 
! -----------------------------------------------------------------------------

! bindCols
! Bind the queue, group or seperate fields to the columns in the result set.
! column order is typically the  same order as the select statment,
! 
! parameters for the ODBC api call 
! hStmt   = handle to the ODBC statement
! colId = ord value of the parmaeter, 1, 2, 3 ... the ordinal position
! colType = the C data type of the column 
! colValue = pointer to the queue field 
! colSize = the size of the buffer or the queue field 
! colInd = pointer to a buffer for the size of the parameter. not used and null in this example 
! -----------------------------------------------------------------------------    
columnsClass.bindColumns procedure(long hStmt) ! sqlReturn

retv      sqlReturn
colInd    &long
x         long,auto
startRow  long,auto

  code 
  
  ! iterate over the list, if any fail return an error
  loop x = 1 to records(self.colq)
    get(self.colQ, x)
    retv = SQLBindCol(hStmt, self.colQ.colId, self.Colq.colType, self.colQ.colValue, self.Colq.colSize, colInd)
    if (retv <> sql_Success) and (retv <> Sql_Success_With_Info) 
      break
    end  
  end   
  
  ! don't care about info messages here
  if (retv = Sql_Success_With_Info)
    retv = Sql_Success
  end
    
  return retv 
! end bindColumns
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! free the queue and the dyn str
! ------------------------------------------------------------------------------
columnsClass.clearQ procedure()

  code 
  
  free(self.colQ)
  self.fieldList.kill()
    
  return
! end clearQ
! ------------------------------------------------------------------------------

! ------------------------------------------------------------------------------
! the various addColumn functions are called by the using code and are used for the
! specific data types.  each calls the AddColumn/3 function to actually 
! add a columns
! ------------------------------------------------------------------------------
columnsClass.AddColumn procedure(*long colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_SLONG, address(colPtr), 4)  
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------
 
columnsClass.AddColumn procedure(string colLabel, *long colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_SLONG, address(colPtr), 4)  
  
  self.addField(colLabel) 
     
  return retv
! end AddColumn
! ------------------------------------------------------------------------------
 
columnsClass.AddColumn procedure(*string colPtr) !,sqlReturn,proc     

retv   sqlReturn(sql_Success)

  code 

  self.addColumn(SQL_C_CHAr, address(colPtr), len(colPtr))  
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(string colLabel, *string colPtr) !,sqlReturn,proc     

retv   sqlReturn(sql_Success)

  code 

  self.addColumn(SQL_C_CHAr, address(colPtr), len(colPtr))  
  self.addField(colLabel) 
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*real colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_DOUBLE, address(colPtr), size(colPtr))
     
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(string colLabel, *real colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_DOUBLE, address(colPtr), size(colPtr))
  self.addField(colLabel) 
     
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(*TIMESTAMP_STRUCT colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_TYPE_TIMESTAMP, address(colPtr), size(TIMESTAMP_STRUCT))
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.AddColumn procedure(string colLabel, *TIMESTAMP_STRUCT colPtr) !,sqlReturn,proc

retv   sqlReturn(sql_Success)

  code 
  
  self.addColumn(SQL_C_TYPE_TIMESTAMP, address(colPtr), size(TIMESTAMP_STRUCT))
  self.addField(colLabel) 
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------
  
! ------------------------------------------------------------------------------
! add a column to the objects queue. 
! the colums will be bound when the execute function is called.
! typically called by the various addColumn functions but can be called directly
! ------------------------------------------------------------------------------  

columnsClass.AddColumn procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength)

retv   sqlReturn(sql_Success)

  code 
  
  ! order is the order the columns are added
  self.colQ.ColId = records(self.colq) + 1
  self.colq.ColType = targetType
  self.colQ.ColValue = TargetValuePtr
  self.colQ.ColSize = BufferLength

  add(self.Colq)
   
  return retv
! end AddColumn
! ------------------------------------------------------------------------------

columnsClass.addField procedure(string colLabel) 

  code 
  
  if (self.fieldList.strLen() > 0)
    self.fieldList.cat(', ' & colLabel)
  else
    self.fieldList.cat(colLabel)
  end
  
  return
! end addField
! ------------------------------------------------------------------------------  