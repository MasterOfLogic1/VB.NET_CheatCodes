# VB.NET_CheatCodes

```markdown
# RPA Code Repository

This repository contains a collection of reusable RPA (Robotic Process Automation) scripts and functions written in VB.Net. Each script addresses a specific task commonly encountered in RPA development, from data manipulation to table joins, date conversions, and more. Below is a detailed description of the scripts:

## Code Descriptions

### 1. **Reading from CSV Using DB**
   Reads data from a CSV file using a database connection.

   ```vb
   Connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+in_processingFolderPath+"; Extended Properties=""text;HDR=NO;FORMAT=Delimited"""
   ```

   - **Sample Query**: Fetches rows based on specific string matches.
   
   ```vb
   "Select [F12] from ["+in_CompilationFileNameWithExtension+"] where  [F12] LIKE '%VNM%' OR [F12] LIKE '%PSM%' OR [F12] LIKE '%NXG%' OR [F12] LIKE '%TAMD%' OR [F12] LIKE '%X00%' OR [F12] LIKE '%SL00%'"
   ```

### 2. **Adding Data Column with Specific Type**
   Adds a new column to a DataTable with a specific data type (e.g., String).

   ```vb
   Trsf_DT.Columns.Add("NARRATION", GetType(String))
   ```

### 3. **Joining Data Tables - Matching Elements**
   Joins two data tables based on matching elements.

   ```vb
   Try
       dtjoin = dtSpoolDT.AsEnumerable().Where(Function(row) dtCompilationDT.AsEnumerable().Any(Function(x) x("NARRATION").ToString = row("NARRATION").ToString)).CopyToDataTable
   Catch ex As Exception
       errorMessage = ex.Message
   End Try
   ```

### 4. **Joining Data Tables - Non-Matching Elements**
   Retrieves rows that do not match between two data tables.

   ```vb
   Try
       dtLeftOver = dtSpoolDT.AsEnumerable().Where(Function(row) Not dtCompilationDT.AsEnumerable().Any(Function(x) x("NARRATION").ToString = row("NARRATION").ToString)).CopyToDataTable
   Catch ex As Exception
       errorMessage = ex.Message
   End Try
   ```

### 5. **Removing White Spaces from Array**
   Removes white spaces from an array or list.

   ```vb
   textArray.Where(Function(x) Not String.IsNullOrWhiteSpace(x)).ToArray
   ```

### 6. **Converting String to DateTime**
   Converts a string into a DateTime object using a specified format.

   ```vb
   Convert.ToDateTime(Datetime.ParseExact(StringRequiredDateTimeStamp, in_ExternalDictionary("DataBaseDateTimeFormat").ToString, System.Globalization.CultureInfo.InvariantCulture))
   ```

### 7. **Removing Element in Array by Name**
   Removes an element from an array by its name.

   ```vb
   io_ArrayToUpdate.Where(Function(s) s <> in_FileToRemove).ToArray
   ```

### 8. **Get All Available Column Indexes**
   Retrieves the index of all columns in a DataTable.

   ```vb
   (From DataColumnFound In out_dtInputData.Columns.Cast(Of DataColumn) Select DataColumnFound.Ordinal).ToList
   ```

### 9. **Merging DataTables**
   Merges two DataTables, copying the parent into the child table.

   ```vb
   Try
       ChildTable.Merge(ParentTable, False, MissingSchemaAction.Ignore)
   Catch ex As Exception
       errorMessage = ex.Message
   End Try
   ```

### 10. **Trimming Spaces in DataTable**
   Removes spaces from all cells in a DataTable.

   ```vb
   TempData = (From r In dtRaw.AsEnumerable Select ia = r.ItemArray.ToList Select ic = ia.ConvertAll(Function(e) e.ToString.Trim).ToArray() Select TempData.Rows.Add(ic)).CopyToDataTable()
   ```

### 11. **Summing Column Values**
   Sums values in a specific column.

   ```vb
   replacedebitTxn.AsEnumerable.Sum(Function(x) Convert.ToDouble(x("TXN_AMT").ToString.Trim)).ToString
   ```

### 12. **Filtering a DataTable**
   Filters a DataTable based on a condition.

   ```vb
   Dim TempDt As System.Data.DataTable = S4spooldt.AsEnumerable().Where(Function(r) r(TradeIdColumnIndex).ToString.StartsWith(TradeID)).CopyToDataTable
   ```

### 13. **Removing Special Characters with Regex**
   Removes special characters from a string.

   ```vb
   System.Text.RegularExpressions.Regex.Replace(variable, “[^a-z A-Z 0-9]”, “”)
   ```

### 14. **Getting Column Sum in DataTable**
   Sums up the values in a column.

   ```vb
   dtResult = (From d In dtData.AsEnumerable Group d By k1 = d(0).ToString, k2 = d(1).ToString.Trim Into grp = Group Let s = grp.Sum(Function(x) CDbl(x(2).ToString.Trim)) Select dtResult.Rows.Add(k1, k2, s)).CopyToDataTable
   ```

### 15. **Removing Empty Spaces in DataTable**
   Removes empty spaces and special characters from all cells in a DataTable.

   ```vb
   dtCorrected = (From r In dt.AsEnumerable Select ia = r.ItemArray.ToList Select ic = ia.ConvertAll(Function(e) System.Text.RegularExpressions.Regex.Replace(e.ToString.Trim.Replace(" ", String.Empty), "[^a-z A-Z 0-9]", String.Empty)).ToArray() Select dtCorrected.Rows.Add(ic)).CopyToDataTable()
   ```

### 16. **Hardcoding a Dictionary**
   Hardcodes a dictionary with string values.

   ```vb
   new Dictionary(Of String, String) From {{"0", "string"}, {"1", "string2"}}
   ```

### 17. **Getting Duplicates in DataTable**
   Retrieves duplicate rows based on specific columns.

   ```vb
   Duplicate = dt.AsEnumerable().
       GroupBy(Function(row) New With {Key .REF = CStr(row("BotUniqueID")), Key .ABS = Math.Abs(CDbl(row("LCY_AMOUNT"))) }).
       Where(Function(Group) Group.Count() > 1).ToList.SelectMany(Function(m) m).CopyToDataTable()
   ```

### 18. **Getting Non-Duplicates in DataTable**
   Retrieves non-duplicate rows based on specific columns.

   ```vb
   NonDuplicates = dt.AsEnumerable().
       GroupBy(Function(row) New With {Key .REF = CStr(row("BotUniqueID")), Key .ABS = Math.Abs(CDbl(row("LCY_AMOUNT"))) }).
       Where(Function(Group) Group.Count() = 1).ToList.SelectMany(Function(m) m).CopyToDataTable()
   ```

### 19. **Removing All Spaces in DataTable**
   Removes all spaces from the cells in a DataTable.

   ```vb
   dtCorrected = (From r In dt.AsEnumerable Select ia = r.ItemArray.ToList Select ic = ia.ConvertAll(Function(e) e.ToString.Trim.Replace(" ", String.Empty)).ToArray() Select dtCorrected.Rows.Add(ic)).CopyToDataTable()
   ```

### 20. **Converting DataRow to Dictionary**
   Converts a DataRow into a dictionary.

   ```vb
   row.Table.Columns.Cast(Of DataColumn)().Zip(row.ItemArray, Function(c, v) New With {.ColumnName = c.ColumnName, .Value = v}).ToDictionary(Function(item) item.ColumnName, Function(item) item.Value)
   ```

```
