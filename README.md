# RefreshWait.xlsm

## VBA/Excel: Wait for Power Query when using RefreshAll

Demonstrates the usage of **Application.CalculateUntilAsyncQueriesDone**,<br />which makes the macro wait for completion of the queries

- Macro/Sub "**WaitforWorksheetQueries**" calls Refresh for the queries _contained in the worksheet_, waits for completion of the queries
- Macro/Sub "**WaitforQueries**" calls Application.RefreshAll, i.e. for _all queries in the workbook_; uses Application.CalculateUntilAsyncQueriesDone
- Macro/Sub "**DontWaitforQueries**" calls Application.RefreshAll, but doesn't wait for completion of the queries
- Macro/Sub "**RefreshAllWait**" refreshes all queries in an alternate way to Application.RefreshAll, waits for completion of the queries<br />and allows for more control about the refreshing cycle, if need

See Excel workbook [RefreshWait.xlsm](./RefreshWait.xlsm)

![sheet1](./img/sheet1.png)

