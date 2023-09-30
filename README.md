# RefreshWait.xlsm

VBA/Excel: Wait for Power Query when using RefreshAll

See exanple file 

Contains two macros and a query.

- Macro "**DontWaitforQueries**" calls Application.RefreshAll, but doesn't wait for completion of the queries
- Macro "**WaitforQueries**" calls Application.RefreshAll followed by Application.CalculateUntilAsyncQueriesDone, which makes the macro wait for completion of the queries
