A small utility script for creating SQL scripts.

In my work, we provide tablets to auditors who go out to polling places on election day and conduct accessibility
audits. Each of those tablets has its own SQL Server instance, containing all of the surveys that auditor is expected
to complete.

Before each election, each of those tablets gets loaded with the surveys for that route for that election. Meaning that
every election day, custom scripts need to be created for every tablet. Previously, this was done by hand.

This script allows the user to select an Excel spreadsheet. That spreadsheet is expected to have column headings
in the first row. The first column is expected to be the route number or name, the second column is the polling place name
and the third column is the unique PPLID that identifies that polling place in the database. Other columns may or may
not contain data. They will be ignored.

The script then, for each unique route number or name, creates a sql script file. Into that file it writes what I refer
to as the "preamble", which cleans up the results, if any, from the last election.

Then it creates an insert command, and for each polling place in that route inserts the polling place id into the route
table. The routes are commented with the polling place name for QA review and to facilitate post facto changes.

Those sql scripts are then individually loaded into the tablets and run with SQL Server.