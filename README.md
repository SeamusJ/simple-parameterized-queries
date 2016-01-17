# simple-parameterized-queries

This is a class for creating super simple parameterized queries in classic ASP. 

Classic ASP is a technology that has not been current in well over a decade. But despite that, there are a lot of legacy applications still floating around written in classic ASP. Should you find yourself tasked with maintaining one of those applications, this class may help to clean up some of that code.

This class is written in JScript, Microsoft's ECMA Script 3 implementation for Classic ASP. This allows the creation of a class with methods that handle an arbitratry number of parameters. And because Classic ASP supports both VB Script and JScript, your VB Script code can create and use instances of this class despite the fact that it is written in a different scripting language.

###Usage

To use the class, include Database.asp:

```
<!--#include file="includes\Database.asp"-->
```

Create an instance of the class:

```
dim database : set database = Database(myConnectionStringHere)
```

The class has three methods: Query, ExecSql, and CreateUpdatableRecordSet.

####Query

The Query method executes a parameterized query, and returns a read-only ADODB Recordset. The query may have any number of parameters.

#####Example

```
set recordSet = database.Query("select user_name from user_table where id=? and status=?", 100, "Active")
```

####ExecSql

The ExecSql method executes parameterized sql. Depending on the sql, this method may return a read-only ADODB Recordset. The query may have any number of parameters.

#####Example

```
database.ExecSql("update user_table set user_name=? where id=?", userName, id)
```

####CreateUpdatableRecordSet

The CreateUpdatableRecordSet is exactly like the Query method except that it returns an updatable ADODB Recordset.

#####Example

```
set recordSet = database.CreateUpdatableRecordSet("select user_name from user_table where id=? and status=?", 100, "Active")
```
