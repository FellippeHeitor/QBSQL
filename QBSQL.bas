CONST verbose = -1

result = createDatabase("testDB", "")
IF NOT result THEN
    result = loadDatabase("testDB", tableCount$)
    IF NOT result THEN PRINT "Error creating/loading database": END
    PRINT "Database loaded; "; tablecount; " tables."
END IF

SLEEP

result = createTable("Persons", "PersonID autoincrement:primary:int,LastName text,FirstName text,Address text,City text")
IF NOT result THEN PRINT "Error creating table 1.": END

SLEEP

result = insertInto("Persons", "LastName,FirstName,City", _
                               "'Heitor','Fellippe','Carandiru';'Soames','Carly','Etoile';'Sinclair','Clair','Carandiru'")
IF NOT result THEN PRINT "Error inserting record(s).": END

SLEEP

result = update("Persons", "Address='43 Fifth Avenue - St. Michel',City='Etoile'", _
                           "")
IF NOT result THEN PRINT "Error updating record(s).": END

SLEEP

result = insertInto("Persons", "", _
                               "'Heitor','Robert','52nd - Saint Louise','Etoile'")
IF NOT result THEN PRINT "Error inserting record(s).": END

SLEEP

result = deleteFrom("Persons", "FirstName='Fellippe'")
IF NOT result THEN PRINT "Error deleting record.": END

SLEEP

result = truncateTable("Persons")
IF NOT result THEN PRINT "Error truncating table 1.": END

SLEEP

result = pack
IF NOT result THEN PRINT "Error packing database": END

SLEEP

result = dropTable("Persons")
IF NOT result THEN PRINT "Error dropping table 1.": END

SLEEP

result = createTable("Persons", "PersonID autoincrement:primary:int,LastName text,FirstName text,Address text,City text")
IF NOT result THEN PRINT "Error creating table 2.": END

SLEEP

result = insertInto("Persons", "", _
                               "'Heitor','Robert','52nd - Saint Louise','Etoile'")
IF NOT result THEN PRINT "Error inserting record(s).": END

SLEEP

result = pack
IF NOT result THEN PRINT "Error packing database": END

FUNCTION deleteFrom%% (this$, __where$)
    'DELETE FROM Customers WHERE CustomerName='Alfreds Futterkiste';

    SHARED currentIniFileName$, IniCODE
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    set$ = __set$
    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF

    DIM columns AS LONG, columnCount AS LONG

    thisTable$ = "table-" + index$
    columnSummary$ = ReadSetting("", thisTable$, "columnsummary")
    columnCount = VAL(ReadSetting("", thisTable$, "columns"))

    'validate where$
    IF LEN(where$) THEN
        REDIM condition$(1 TO columnCount, 1 TO 3)
        columns = 0
        arg$ = where$
        DO
            columns = columns + 1
            c = INSTR(arg$, ",")
            DO
                IF c > 1 THEN
                    IF ASC(arg$, c - 1) = 92 THEN
                        c = INSTR(c + 1, arg$, ",")
                    ELSE
                        EXIT DO
                    END IF
                ELSE
                    EXIT DO
                END IF
            LOOP
            IF c THEN
                thisArg$ = LEFT$(arg$, c - 1)
                arg$ = MID$(arg$, c + 1)
            ELSE
                thisArg$ = arg$
            END IF

            eq = INSTR(thisArg$, "=")
            IF eq = 0 THEN
                eq = INSTR(thisArg$, "<>")
                IF eq = 0 THEN
                    eq = INSTR(thisArg$, "<")
                    IF eq = 0 THEN
                        eq = INSTR(thisArg$, ">")
                        IF eq = 0 THEN
                            'syntax error
                            IF verbose THEN PRINT "Condition requires =, <>, < or > operators in WHERE clause."
                            EXIT FUNCTION
                        ELSE
                            condition$(columns, 2) = ">"
                        END IF
                    ELSE
                        condition$(columns, 2) = "<"
                    END IF
                ELSE
                    condition$(columns, 2) = "<>"
                END IF
            ELSE
                condition$(columns, 2) = "="
            END IF

            thisValue$ = MID$(thisArg$, eq + 1)
            thisArg$ = LEFT$(thisArg$, eq - 1)

            IF LEFT$(thisValue$, 1) = "'" AND RIGHT$(thisValue$, 1) = "'" THEN
                thisValue$ = MID$(thisValue$, 2, LEN(thisValue$) - 2)
            END IF

            condition$(columns, 1) = thisArg$
            condition$(columns, 3) = thisValue$
            IF INSTR(columnSummary$, "/" + thisArg$ + "/") = 0 THEN
                'this column does not exist
                IF verbose THEN PRINT "Column '" + thisArg$ + "' does not exist in table '" + this$ + "'."
                EXIT FUNCTION
            END IF
        LOOP WHILE c > 0
        totalConditions = columns
    END IF

    'perform the deletion; go through whole table
    totalRecords& = 0
    FOR i& = 1 TO recordCount(this$)
        IF ReadSetting("", thisTable$ + "-" + LTRIM$(STR$(i&)), "state") = "deleted" THEN _CONTINUE
        conditionsMet = -1
        FOR check& = 1 TO totalConditions
            thisValue$ = ReadSetting("", thisTable$ + "-" + LTRIM$(STR$(i&)), condition$(check&, 1))
            SELECT CASE condition$(check&, 2)
                CASE "="
                    IF thisValue$ <> condition$(check&, 3) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE "<>"
                    IF thisValue$ = condition$(check&, 3) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE ">"
                    IF VAL(thisValue$) <= VAL(condition$(check&, 3)) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE "<"
                    IF VAL(thisValue$) >= VAL(condition$(check&, 3)) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
            END SELECT
        NEXT

        IF conditionsMet THEN
            'mark this record for deletion; decrease record count
            totalRecords& = totalRecords& + 1
            FOR commit& = 1 TO totalItemsInSet
                WriteSetting "", thisTable$ + "-" + LTRIM$(STR$(i&)), "state", "deleted"
            NEXT
        END IF
    NEXT

    deleteFrom%% = -1
    IF verbose THEN PRINT totalRecords&; "records marked for deletion."
END FUNCTION

FUNCTION update%% (this$, __set$, where$)
    'UPDATE Customers
    'SET ContactName = 'Alfred Schmidt', City= 'Frankfurt'
    'WHERE CustomerID = 1;

    SHARED currentIniFileName$, IniCODE
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    set$ = __set$
    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF
    IF LEN(set$) = 0 THEN
        IF verbose THEN PRINT "Invalid syntax: SET not passed."
        EXIT FUNCTION
    END IF

    DIM columns AS LONG, columnCount AS LONG

    thisTable$ = "table-" + index$
    columnSummary$ = ReadSetting("", thisTable$, "columnsummary")
    columnCount = VAL(ReadSetting("", thisTable$, "columns"))

    'validate set$
    REDIM dataSet$(1 TO columnCount, 1 TO 2)
    columns = 0
    arg$ = set$
    DO
        columns = columns + 1
        c = INSTR(arg$, ",")
        DO
            IF c > 1 THEN
                IF ASC(arg$, c - 1) = 92 THEN
                    c = INSTR(c + 1, arg$, ",")
                ELSE
                    EXIT DO
                END IF
            ELSE
                EXIT DO
            END IF
        LOOP
        IF c THEN
            thisArg$ = LEFT$(arg$, c - 1)
            arg$ = MID$(arg$, c + 1)
        ELSE
            thisArg$ = arg$
        END IF

        eq = INSTR(thisArg$, "=")
        IF eq = 0 THEN
            'syntax error
            IF verbose THEN PRINT "Invalid syntax; expected key=value format in SET."
            EXIT FUNCTION
        END IF

        thisValue$ = MID$(thisArg$, eq + 1)
        thisArg$ = LEFT$(thisArg$, eq - 1)

        IF LEFT$(thisValue$, 1) = "'" AND RIGHT$(thisValue$, 1) = "'" THEN
            thisValue$ = MID$(thisValue$, 2, LEN(thisValue$) - 2)
        END IF

        dataSet$(columns, 1) = thisArg$
        dataSet$(columns, 2) = Replace(thisValue$, "\,", ",", 0, 0)
        IF INSTR(columnSummary$, "/" + thisArg$ + "/") = 0 THEN
            'this column does not exist
            IF verbose THEN PRINT "Column '" + thisArg$ + "' does not exist in table '" + this$ + "'."
            EXIT FUNCTION
        END IF

        'scan for autoincrement columns
        FOR i& = 1 TO columnCount
            check$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-autoincrement")
            IF IniCODE = 0 THEN
                'key found; this column must be added with an autoincrement value
                thisColumn$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-name")
                IF thisColumn$ = thisArg$ THEN
                    'can't directly specify the value of an autoincrement column
                    IF verbose THEN PRINT "Cannot set column '" + thisArg$ + "' as it is set to autoincrement in table '" + this$ + "'."
                    EXIT FUNCTION
                END IF
            END IF
        NEXT
    LOOP WHILE c > 0
    totalItemsInSet = columns

    'validate where$
    IF LEN(where$) THEN
        REDIM condition$(1 TO columnCount, 1 TO 3)
        columns = 0
        arg$ = where$
        DO
            columns = columns + 1
            c = INSTR(arg$, ",")
            DO
                IF c > 1 THEN
                    IF ASC(arg$, c - 1) = 92 THEN
                        c = INSTR(c + 1, arg$, ",")
                    ELSE
                        EXIT DO
                    END IF
                ELSE
                    EXIT DO
                END IF
            LOOP
            IF c THEN
                thisArg$ = LEFT$(arg$, c - 1)
                arg$ = MID$(arg$, c + 1)
            ELSE
                thisArg$ = arg$
            END IF

            eq = INSTR(thisArg$, "=")
            IF eq = 0 THEN
                eq = INSTR(thisArg$, "<>")
                IF eq = 0 THEN
                    eq = INSTR(thisArg$, "<")
                    IF eq = 0 THEN
                        eq = INSTR(thisArg$, ">")
                        IF eq = 0 THEN
                            'syntax error
                            IF verbose THEN PRINT "Condition requires =, <>, < or > operators in WHERE clause."
                            EXIT FUNCTION
                        ELSE
                            condition$(columns, 2) = ">"
                        END IF
                    ELSE
                        condition$(columns, 2) = "<"
                    END IF
                ELSE
                    condition$(columns, 2) = "<>"
                END IF
            ELSE
                condition$(columns, 2) = "="
            END IF

            thisValue$ = MID$(thisArg$, eq + 1)
            thisArg$ = LEFT$(thisArg$, eq - 1)

            IF LEFT$(thisValue$, 1) = "'" AND RIGHT$(thisValue$, 1) = "'" THEN
                thisValue$ = MID$(thisValue$, 2, LEN(thisValue$) - 2)
            END IF

            condition$(columns, 1) = thisArg$
            condition$(columns, 3) = thisValue$
            IF INSTR(columnSummary$, "/" + thisArg$ + "/") = 0 THEN
                'this column does not exist
                IF verbose THEN PRINT "Column '" + thisArg$ + "' does not exist in table '" + this$ + "'."
                EXIT FUNCTION
            END IF
        LOOP WHILE c > 0
        totalConditions = columns
    END IF

    'perform the update; go through whole table
    totalRecords& = 0
    FOR i& = 1 TO recordCount(this$)
        IF ReadSetting("", thisTable$ + "-" + LTRIM$(STR$(i&)), "state") = "deleted" THEN _CONTINUE
        conditionsMet = -1
        FOR check& = 1 TO totalConditions
            thisValue$ = ReadSetting("", thisTable$ + "-" + LTRIM$(STR$(i&)), condition$(check&, 1))
            SELECT CASE condition$(check&, 2)
                CASE "="
                    IF thisValue$ <> condition$(check&, 3) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE "<>"
                    IF thisValue$ = condition$(check&, 3) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE ">"
                    IF VAL(thisValue$) <= VAL(condition$(check&, 3)) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
                CASE "<"
                    IF VAL(thisValue$) >= VAL(condition$(check&, 3)) THEN
                        conditionsMet = 0
                        EXIT FOR
                    END IF
            END SELECT
        NEXT

        IF conditionsMet THEN
            'update this record
            totalRecords& = totalRecords& + 1
            FOR commit& = 1 TO totalItemsInSet
                WriteSetting "", thisTable$ + "-" + LTRIM$(STR$(i&)), dataSet$(commit&, 1), dataSet$(commit&, 2)
            NEXT
        END IF
    NEXT

    update%% = -1
    IF verbose THEN PRINT totalRecords&; "records affected."
END FUNCTION

FUNCTION insertInto%% (this$, columns$, __values$)
    'INSERT INTO Customers (CustomerName, ContactName, Address, City, PostalCode, Country)
    'VALUES ('Cardinal', 'Tom B. Erichsen', 'Skagen 21', 'Stavanger', '4006', 'Norway');

    SHARED currentIniFileName$, IniCODE
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    values$ = __values$
    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF
    IF LEN(values$) = 0 THEN
        IF verbose THEN PRINT "No values passed."
        EXIT FUNCTION
    END IF

    DIM columns AS LONG, columnCount AS LONG

    thisTable$ = "table-" + index$
    columnSummary$ = ReadSetting("", thisTable$, "columnsummary")
    columnCount = VAL(ReadSetting("", thisTable$, "columns"))

    'validate columns$
    REDIM columnName$(1 TO columnCount), columnType$(1 TO columnCount)
    columns = 0
    totalColumns = 0
    arg$ = columns$
    DO
        columns = columns + 1

        IF LEN(columns$) > 0 THEN
            c = INSTR(arg$, ",")
            DO
                IF c > 1 THEN
                    IF ASC(arg$, c - 1) = 92 THEN
                        c = INSTR(c + 1, arg$, ",")
                    ELSE
                        EXIT DO
                    END IF
                ELSE
                    EXIT DO
                END IF
            LOOP

            IF c THEN
                thisArg$ = LEFT$(arg$, c - 1)
                arg$ = MID$(arg$, c + 1)
            ELSE
                thisArg$ = arg$
            END IF

            totalColumns = totalColumns + 1
            columnName$(totalColumns) = thisArg$
            IF INSTR(columnSummary$, "/" + thisArg$ + "/") = 0 THEN
                'this column does not exist
                IF verbose THEN PRINT "Column '" + thisArg$ + "' doesn't exist in table '" + this$ + "'."
                EXIT FUNCTION
            END IF

            'scan for types and for autoincrement columns
            FOR i& = 1 TO columnCount
                checkName$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-name")
                IF checkName$ = thisArg$ THEN
                    columnType$(totalColumns) = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-type")
                END IF

                check$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-autoincrement")
                IF IniCODE = 0 THEN
                    'key found; this column must be added with an autoincrement value
                    thisColumn$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-name")
                    IF thisColumn$ = thisArg$ THEN
                        'can't directly specify the value of an autoincrement column
                        IF verbose THEN PRINT "Cannot set column '" + thisArg$ + "' as it is set to autoincrement in table '" + this$ + "'."
                        EXIT FUNCTION
                    END IF
                END IF
            NEXT
        ELSE
            c = 1
            IF columns > columnCount THEN EXIT DO
            check$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(columns))) + "-autoincrement")
            IF IniCODE = 0 THEN _CONTINUE
            totalColumns = totalColumns + 1
            columnName$(totalColumns) = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(columns))) + "-name")
            columnType$(totalColumns) = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(columns))) + "-type")
        END IF

    LOOP WHILE c > 0
    REDIM _PRESERVE columnName$(1 TO totalColumns)

    'insert records
    totalRecords& = 0
    DO
        nextRecord$ = ""
        columns = 0
        DO
            c = INSTR(values$, ",")
            DO
                IF c > 1 THEN
                    IF ASC(values$, c - 1) = 92 THEN
                        c = INSTR(c + 1, values$, ",")
                    ELSE
                        EXIT DO
                    END IF
                ELSE
                    EXIT DO
                END IF
            LOOP
            d = INSTR(values$, ";")
            IF c THEN
                IF d > 0 AND d < c THEN c = d
                thisValue$ = LEFT$(values$, c - 1)
                values$ = MID$(values$, c + 1)
                IF LEFT$(values$, 1) = ";" THEN
                    'new record
                    values$ = MID$(values$, 2)
                    EXIT DO
                END IF
            ELSE
                thisValue$ = values$
            END IF

            IF LEFT$(thisValue$, 1) = "'" AND RIGHT$(thisValue$, 1) = "'" THEN
                thisValue$ = MID$(thisValue$, 2, LEN(thisValue$) - 2)
            END IF

            thisValue$ = Replace(thisValue$, "\,", ",", 0, 0)

            columns = columns + 1
            IF columns > UBOUND(columnName$) THEN
                'invalid dataset passed
                IF verbose THEN PRINT "SET: invalid number of columns passed."
                EXIT FUNCTION
            END IF

            IF nextRecord$ = "" THEN
                totalRecords& = totalRecords& + 1
                nextRecord$ = LTRIM$(STR$(VAL(ReadSetting("", thisTable$, "recordindex")) + 1))
                WriteSetting "", thisTable$, "recordindex", nextRecord$
                WriteSetting "", thisTable$, "recordcount", STR$(VAL(ReadSetting("", thisTable$, "recordcount")) + 1)

                WriteSetting "", thisTable$ + "-" + nextRecord$, "state", "active"

                'scan for autoincrement columns
                FOR i& = 1 TO columnCount
                    check$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-autoincrement")
                    IF IniCODE = 0 THEN
                        'key found; this column must be added with an autoincrement value
                        thisColumn$ = ReadSetting("", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-name")
                        check$ = STR$(VAL(check$) + 1)
                        WriteSetting "", thisTable$, "column-" + LTRIM$(RTRIM$(STR$(i&))) + "-autoincrement", check$
                        WriteSetting "", thisTable$ + "-" + nextRecord$, thisColumn$, check$
                    END IF
                NEXT
            END IF

            thisValue$ = validateValue(thisValue$, columnType$(columns))
            WriteSetting "", thisTable$ + "-" + nextRecord$, columnName$(columns), thisValue$
            IF c = d THEN EXIT DO
        LOOP WHILE c > 0
        IF c = 0 THEN EXIT DO
    LOOP

    insertInto%% = -1
    IF verbose THEN PRINT totalRecords&; "records inserted."
END FUNCTION

FUNCTION validateValue$ (__value$, __dataType$)
    datatype$ = __dataType$

    bracket = INSTR(datatype$, "(")
    IF bracket THEN
        length& = VAL(MID$(datatype$, bracket + 1))
        datatype$ = LEFT$(datatype$, bracket - 1)
    END IF

    SELECT CASE UCASE$(datatype$)
        CASE "INT"
            validateValue$ = LTRIM$(STR$(INT(VAL(__value$))))
            EXIT FUNCTION
        CASE "FLOAT"
            validateValue$ = LTRIM$(STR$(VAL(__value$)))
            EXIT FUNCTION
        CASE "TEXT"
            IF length& THEN
                validateValue$ = LEFT$(__value$, length&)
                EXIT FUNCTION
            END IF
    END SELECT
    validateValue$ = __value$
END FUNCTION

FUNCTION recordCount& (this$)
    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF

    recordCount& = VAL(ReadSetting("", "table-" + index$, "recordcount"))
END FUNCTION

FUNCTION dropTable%% (this$)
    SHARED currentIniFileName$
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF

    result = truncateTable(this$)

    thisTable$ = "table-" + index$
    WriteSetting "", thisTable$, "state", "deleted"

    'decrease total table count
    WriteSetting "", "database", "tablecount", STR$(VAL(ReadSetting("", "database", "tablecount")) - 1)

    dropTable%% = -1
    IF verbose THEN PRINT "Table '" + this$ + "' dropped."
END FUNCTION

FUNCTION truncateTable%% (this$)
    SHARED currentIniFileName$, IniCODE
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    IF tableExists(this$, index$) = 0 THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF

    'mark all records for deletion:
    totalRecords& = recordCount(this$)
    FOR i& = 1 TO totalRecords&
        WriteSetting "", "table-" + index$ + "-" + LTRIM$(STR$(i&)), "state", "deleted"
    NEXT

    WriteSetting "", "table-" + index$, "recordcount", "0"

    DIM columnCount AS LONG
    columnCount = VAL(ReadSetting("", "table-" + index$, "columns"))
    FOR i& = 1 TO columnCount
        check$ = ReadSetting("", "table-" + index$, "column-" + LTRIM$(STR$(i&)) + "-autoincrement")
        IF IniCODE = 0 THEN
            'key found; reset its value
            WriteSetting "", "table-" + index$, "column-" + LTRIM$(STR$(i&)) + "-autoincrement", "0"
        END IF
    NEXT

    truncateTable%% = -1
    IF verbose THEN PRINT "Table '" + this$ + "' truncated;"; totalRecords&; "marked for deletion."
END FUNCTION

FUNCTION pack%%
    SHARED currentIniFileName$, IniCODE
    SHARED IniLastSection$, IniLastKey$

    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    'Read a key from [database] to set pointer to top of file
    check$ = ReadSetting("", "database", "name")

    DO
        check$ = ReadSetting$("", "", "")
        IF IniCODE THEN EXIT DO
        IF IniLastKey$ = "state" THEN
            check$ = ReadSetting$("", IniLastSection$, "state")
            IF check$ = "deleted" THEN
                IniDeleteSection "", IniLastSection$
            END IF
        END IF
    LOOP

    pack%% = -1
    IF verbose THEN PRINT "Database successfully purged."
END FUNCTION

FUNCTION tableExists%% (this$, index$)
    FOR i& = 1 TO VAL(ReadSetting("", "database", "tableindex"))
        check$ = UCASE$(ReadSetting("", "tables", STR$(i&)))
        IF UCASE$(this$) = check$ THEN
            IF ReadSetting("", "table-" + LTRIM$(STR$(i&)), "state") = "active" THEN
                'table exists and is active
                index$ = LTRIM$(STR$(i&))
                tableExists%% = -1
                EXIT FUNCTION
            END IF
        END IF
    NEXT
END FUNCTION

FUNCTION createTable%% (this$, arg$)
    'CREATE TABLE Persons (
    '    PersonID int,
    '    LastName varchar(255),
    '    FirstName varchar(255),
    '    Address varchar(255),
    '    City varchar(255)
    ');

    SHARED currentIniFileName$
    IF currentIniFileName$ = "" THEN
        IF verbose THEN PRINT "No database loaded."
        EXIT FUNCTION
    END IF

    IF tableExists(this$, index$) THEN
        IF verbose THEN PRINT "Table '" + this$ + "' does not exist"
        EXIT FUNCTION
    END IF

    nextId$ = LTRIM$(STR$(VAL(ReadSetting("", "database", "tableindex")) + 1))

    'increase tableindex and total table count
    WriteSetting "", "database", "tableindex", nextId$
    WriteSetting "", "database", "tablecount", STR$(VAL(ReadSetting("", "database", "tablecount")) + 1)

    'create table entry
    WriteSetting "", "tables", nextId$, this$

    'create table
    newTable$ = "table-" + nextId$
    WriteSetting "", newTable$, "date", DATE$
    WriteSetting "", newTable$, "time", TIME$
    WriteSetting "", newTable$, "state", "active"
    WriteSetting "", newTable$, "columns", "0"
    WriteSetting "", newTable$, "columnsummary", "//"
    WriteSetting "", newTable$, "recordcount", "0"
    WriteSetting "", newTable$, "recordindex", "0"

    IF LEN(arg$) = 0 THEN
        IF verbose THEN PRINT "Table structure not passed."
        EXIT FUNCTION
    END IF

    IF INSTR(LCASE$(arg$), "primary:") THEN
        WriteSetting "", newTable$, "primarykey", "0"
        WriteSetting "", newTable$, "keyindex", "0"
    END IF

    columns = 0
    columnSummary$ = "/"
    DO
        columns = columns + 1

        c = INSTR(arg$, ",")
        IF c THEN
            thisArg$ = LEFT$(arg$, c - 1)
            arg$ = MID$(arg$, c + 1)
        ELSE
            thisArg$ = arg$
        END IF

        sp = INSTR(thisArg$, " ")
        thisColumn$ = LEFT$(thisArg$, sp - 1)

        IF INSTR(columnSummary$, "/" + thisColumn$ + "/") > 0 THEN
            'duplicated column definition
            IF verbose THEN PRINT "Column '" + thisColumn$ + "' defined more than once."
            EXIT FUNCTION
        END IF
        columnSummary$ = columnSummary$ + thisColumn$ + "/"

        thisType$ = MID$(thisArg$, sp + 1)
        WriteSetting "", newTable$, "column-" + LTRIM$(STR$(columns)) + "-name", thisColumn$

        IF checkConstraint(thisType$, "primary") THEN
            WriteSetting "", newTable$, "primarykey", STR$(columns)
        END IF

        IF checkConstraint(thisType$, "autoincrement") THEN
            WriteSetting "", newTable$, "column-" + LTRIM$(STR$(columns)) + "-autoincrement", "0"
            IF thisType$ <> "int" THEN
                IF verbose THEN PRINT "Autoincrement columns must be of type INT."
                EXIT FUNCTION
            END IF
        END IF

        WriteSetting "", newTable$, "column-" + LTRIM$(STR$(columns)) + "-type", thisType$
    LOOP WHILE c > 0

    WriteSetting "", newTable$, "columns", STR$(columns)
    WriteSetting "", newTable$, "columnsummary", columnSummary$

    createTable%% = -1
    IF verbose THEN PRINT "Table '" + this$ + "' created;"; columns; "columns defined."
END FUNCTION

FUNCTION checkConstraint%% (__this$, __constraint$)
    constraint$ = UCASE$(__constraint$)
    this$ = UCASE$(__this$)

    found = INSTR(this$, constraint$ + ":")
    IF found THEN
        __this$ = LEFT$(__this$, found - 1) + MID$(__this$, found + LEN(constraint$) + 1)
        checkConstraint%% = -1
    END IF
END FUNCTION

FUNCTION loadDatabase%% (this$, tableCount$)
    fileName$ = toFile(this$)

    IF _FILEEXISTS(fileName$) THEN
        tableCount$ = ReadSetting(fileName$, "database", "tablecount")
        loadDatabase%% = -1
        IF verbose THEN PRINT "Database loaded."
    END IF
END FUNCTION

FUNCTION createDatabase%% (this$, arg$)
    fileName$ = toFile(this$)

    IF _FILEEXISTS(fileName$) THEN
        IF UCASE$(arg$) = "IF NOT EXISTS" THEN
            IF verbose THEN PRINT "Database already exists."
            EXIT FUNCTION
        END IF
        KILL fileName$
    END IF

    WriteSetting fileName$, "database", "date", DATE$
    WriteSetting "", "database", "time", TIME$
    WriteSetting "", "database", "name", this$
    WriteSetting "", "database", "tablecount", "0"
    WriteSetting "", "database", "tableindex", "0"
    createDatabase%% = -1
    IF verbose THEN PRINT "Database successfully created."
END FUNCTION

FUNCTION dropDatabase%% (this$, arg$)
    fileName$ = toFile(this$)

    IF _FILEEXISTS(fileName$) THEN
        KILL fileName$
    END IF

    IniClose

    dropDatabase%% = -1
    IF verbose THEN PRINT "Database dropped."
END FUNCTION

FUNCTION backupDatabase%% (this$, arg$)
    fileName$ = toFile(this$)

    IF _FILEEXISTS(fileName$) = 0 THEN
        IF verbose THEN PRINT "Database '" + this$ + "' does not exist."
        EXIT FUNCTION
    END IF
    IF _FILEEXISTS(arg$) THEN KILL arg$

    f1 = FREEFILE
    OPEN fileName$ FOR BINARY AS #f1
    f2 = FREEFILE
    OPEN arg$ FOR BINARY AS #f2

    a$ = SPACE$(LOF(f1))
    GET #f1, 1, a$
    PUT #f2, 1, a$

    CLOSE f1, f2
    backupDatabase%% = -1
    IF verbose THEN PRINT "Database backup successfully created."
END FUNCTION

FUNCTION toFile$ (__this$)
    this$ = __this$
    IF LCASE$(RIGHT$(this$, 4)) <> ".ini" THEN
        this$ = this$ + ".ini"
    END IF
    toFile$ = this$
END FUNCTION

FUNCTION Replace$ (TempText$, SubString$, NewString$, CaseSensitive AS _BYTE, TotalReplacements AS LONG)
    DIM FindSubString AS LONG, Text$

    IF LEN(TempText$) = 0 THEN EXIT SUB

    Text$ = TempText$
    TotalReplacements = 0
    DO
        IF CaseSensitive THEN
            FindSubString = INSTR(FindSubString + 1, Text$, SubString$)
        ELSE
            FindSubString = INSTR(FindSubString + 1, UCASE$(Text$), UCASE$(SubString$))
        END IF
        IF FindSubString = 0 THEN EXIT DO
        IF LEFT$(SubString$, 1) = "\" THEN 'Escape sequence
            'Replace the Substring if it's not preceeded by another backslash
            IF MID$(Text$, FindSubString - 1, 1) <> "\" THEN
                Text$ = LEFT$(Text$, FindSubString - 1) + NewString$ + MID$(Text$, FindSubString + LEN(SubString$))
                TotalReplacements = TotalReplacements + 1
            END IF
        ELSE
            Text$ = LEFT$(Text$, FindSubString - 1) + NewString$ + MID$(Text$, FindSubString + LEN(SubString$))
            TotalReplacements = TotalReplacements + 1
        END IF
    LOOP

    Replace$ = Text$
END FUNCTION

'$include:'../INI-Manager/ini.bm'
