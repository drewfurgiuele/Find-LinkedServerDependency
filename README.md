# Find-LinkedServerDependantObjects
Ever wonder how many objects have code that reference linked servers in a database? Well, I did. And here's what I wrote to help find them.

## Hey, look, I know linked servers can suck
But they're a neccessary evil sometimes. Chances are you either inheritied a bunch of databases that contain code that references them, or maybe you set some up on purpose. And that's okay! The problem is: how do you know *exactly* what is referencing your linked servers?

## You could manually look at each object definition or do a find/replace
And that might work, but you still might miss things. And what about if you have a linked server that matches a schema or other object names? We're human, and we make mistakes. Why don't we automate our brains?

## That's right: you can parse each object that might contain a linked server reference
You can point this function to a given instance and database name it'll scan:
* Table triggers
* Database synonyms
* Views
* Functions
* Stored procedures

This will parse each object's DDL and look for four-part identifiers and mark those objects are referencing a linked server. The function will return an array of objects that details the referecing object as well as the linked server, database, schema, and object name being called via the linked server.

## Requirements
To use this function, you'll need:
1. The sqlserver PowerShell module (https://www.powershellgallery.com/packages/SqlServer)
2. Access to the Microsoft.SqlServer.TransactSql.ScriptDom assembly. If you run this on a host with a recent(ish) version of SQL Server installed, this won't be an issue. Otherwise, you need to supply a path to this library

## Known issues / todo
1. Linked servers might lurk elsewhere? I don't know.
2. Threading?
