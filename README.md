# GeneralSQLReporter
.dll for General SQL Reporting, provides utilities to run any SELECT query against SQL and generate various outputs

## Setup
Once the NuGet Package / .dll has been added to your Project, you must Setup the `SqlRepository` by calling `SqlRepository.SetupSqlConnection()` like below:
```
var connected = SqlRepository.SetupSqlConnection(connectionString);
```

This uses the default `ConnectionString` defined at https://learn.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlconnection.connectionstring?view=dotnet-plat-ext-7.0

## Creating a `SqlReport`
To create a `SqlReport` that we can run, just create a new instance of the `SqlReport` class like so:
```
var sqlReport = new SqlReport("SELECT * FROM [AvailableCars]", ReportFormat.Csv);
```

With the first parameter in this example being just a standard `SELECT * FROM X` query and the second parameter being the desired Report Output format.

## Creating a `StoredProcedureReport`
Similar to the `SqlReport` you just need to create the instance of the `StoredProcedureReport` class like so:
```
var storedProcedureReport = new StoredProcedureReport("sp_GetCarsByName", ReportFormat.Csv);
```

With the first parameter being the name of the Stored Procedure you wish to run.

## Adding Parameters to a Report (`SqlReport` or `StoredProcedureReport`)
Both the `SqlReport` and `StoredProcedureReport` support using `SqlParameters` (https://learn.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlparameter?view=dotnet-plat-ext-7.0).

In the example I'll be following from the section `Creating a StoredProcedureReport`.

There is a helpful `AddParameter` method on both the `SqlReport` and `StoredProcedureReport` base class (`GenericReport`) which handles defining and assigning the parameters with the values you specify like below:
```
storedProcedureReport.AddParameter("@name", SqlDbType.NVarChar, "Suzuki");
storedProcedureReport.AddParameter("@make", SqlDbType.NVarChar, "Swift");
```

Where the first parameter is the `name` of the `SqlParameter`, the second being the `DbType` / Data Type for the parameter and the last parameter being the value you want to assign.

## Running a Report (`SqlReport` or `StoredProcedureReport`)
As both the `SqlReport` and `StoredProcedureReport` share the same base class (`GenericReport)` the `SqlRepository` can dynamically handle either of these classes.

To run the Report all you need to add is:
```
var storedProcedureResult = SqlRepository.RunSingleReport(storedProcedureReport);
```

With the parameter being passed in (in this example the `storedProcedureReport` created in `Creating a StoredProcedureReport` and the resulting `ReportResultSet` being stored into the `storedProcedureResult` variable for exporting.
