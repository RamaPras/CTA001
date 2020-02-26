Copyright (c) 2005-2020 Teradata Corporation. All rights reserved.

--------------------------------------------------------------------------------
                     .NET Data Provider for Teradata               
                              Version 16.20.9
--------------------------------------------------------------------------------

********************************************************************************
****  NOTE: The Data Provider requires .NET FRAMEWORK 4.5.2 or higher,      ****
****        or .NET CORE 2.0 or higher.                                     ****
****        It does not support .NET Framework 3.5, 4.0, 4.5 or 4.5.1       ****
********************************************************************************

********************************************************************************
****  NOTE: The connection string attribute UseEnhancedSchemaTable has      ****
****        been deprecated. It may be removed in a future release.         ****
********************************************************************************


RELEASE NOTES


Content:

1.0 Overview
1.1   New Features
1.2   Problems Fixed
2.0 Teradata Product Dependencies
3.0 Operating Environments
4.0 Installation
4.1   Installation Procedure
4.2   Installation Options
5.0 Uninstall
6.0 Restrictions and Known Issues



1.0 Overview
============
The .NET Data Provider for Teradata is an implementation of the Microsoft 
ADO.NET specification. It provides direct access to the Teradata Vantage 
and integrates with the DataSet. .NET Applications use the Data Provider
to load data into or retrieve data from the Advanced SQL Engine.

The .NET Data Provider for Teradata is available for download as a NuGet 
package (from https://www.nuget.org/packages/Teradata.Client.Provider) and a
MSI package (from https://downloads.teradata.com).

The MSI package installs the .NET Data Provider for Teradata (ADO.NET) and 
the Entity Framework Provider for Teradata with support for the Entity Data 
Model (EDM), LINQ-to-Entities and Entity-SQL. The Entity Framework Provider 
for Teradata supports ADO.NET Entity Framework 3.5 SP1 features. It does not
support ADO.NET Entity Framework 4.0 (or higher) features and it is not 
compatible with Entity Framework 6.x or Entity Framework Core. The assemblies
included in the MSI package are .NET Framework assemblies. The package does 
not include the assemblies for .NET Core.

The NuGet package contains the Data Provider for .NET Framework and .NET Core.
This package does not include the Entity Framework Provider or the Entity 
Framework Core Provider. 

The Entity Framework Core Provider for Teradata is available as a NuGet package
(from https://www.nuget.org/packages/Teradata.EntityFrameworkCore).

See following links for additional information:
 a) ADO.NET Overview: 
        https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/ado-net-overview

 b) Entity Framework: 
        https://docs.microsoft.com/en-us/ef/



1.1 New Features
================

Version 16.20.9:
  a) The Data Provider closes idle connections, across all pools, before the
     process exits on .NET Core. .NET Core, unlike .NET Framework, does not 
     invoke Finalizers during Process-Exit. Therefore a new scheme had to be 
     introduced to close the connections to the Advanced SQL Engine on .NET 
     Core.

Version 16.20.5:
  a) Support for GetFieldValue<SByte> has been added for ByteInt columns. [NET-4486]

  b) The Data Provider now applies the Connection-Timeout to the Change-Database 
     command. Previously it applied the Default Command-timeout to the 
     Change-Database command when it opened the Connection (TdConnection.Open()). 
     This change will resolve an issues related to the Teradata Workload Analyzer 
     delaying the Change-Database command. An Application must increase the 
     "Connection Timeout" when the Teradata Workload Analyzer is setup to delay 
     the Change-Database command. [NET-4485]

Version 16.20.4:
  a) The Data Provider supports SPNEGO Authentication Mechanism on .NET Core 
     in addition to LDAP and TD2 mechanisms previously supported in 
     version 16.20.2. SPNEGO authentication on .NET Core requires the 
     Teradata Database version 16.10.03 or 16.20.22 and higher. 

  b) The Data Provider supports Direct Access and Materialized Foreign tables.
     The TABLE_TYPE_EX column of the Tables Schema Collection returns
     information for the foreign tables. [NET-4422]

Version 16.20.2:
  a) The .NET Data Provider for Teradata has been ported to .NET Core. [NET-4181]
     Some features are not supported by the .NET Core implementation.
     Please refer to the Migration Considerations topic in the help file.

  b) The tracing functionality now uses Event Tracing for Windows (ETW) on
     Windows platforms and lttng on Linux. [NET-4179]
     Please refer to the Troubleshooting topic in the help file for more details.

  c) The Data Provider supports the JWT authentication mechanism. Refer to
     help topic 'Authentication String Construction' for more details. [NET-4718]

  d) Support for the new Set Transform statement has been added.
     You must use this statement BEFORE preparing a statement. When you issue a
     Set Transform statement any statements that have already been prepared on
     this session will be reset to an un-prepared state. [NET-4049]   

  e) The .NET Data Provider for Teradata supports the JWT authentication 
     mechanism. [NET-4178]

Version 16.20:
  a) A new generic method has been added to TdParameter to retrieve the value of
     an Output or InputOutput parameter as a specified data type. [NET-3732, 
     NET-3996]
        TdParameter.GetValue<T>

  b) The following methods may now be called when retrieving LOB data when using
     CommandBehavior.SequentialAccess. Previously, these methods were only 
     available when using CommandBehavior.Default: [NET-3969]
        TdDataReader.GetStream
        TdDataReader.GetTextReader
        TdDataReader.GetXmlReader
        TdDataReader.GetFieldValue<T>        for T as Stream, TextReader or XmlReader.
        TdDataReader.GetFieldValueAsync<T>   for T as Stream, TextReader or XmlReader.

  c) Support has been added for the new DATASET STORAGE FORMAT CSV data type in
     the Columns Schema Collection and TdDataReader.GetSchemaTable output. 
     [NET-3896]

  d) Support has been added for TIME SERIES tables in the Schema Collections: 
     [NET-3581]
        - The TABLE_TYPE_EX column in the Tables collection may now return TIME
          SERIES, either as the type or as part of the Global Temporary type.
        - The COLUMN_INFO column in the Columns collection may now return TC 
          or TN.

Version 16.10:
  a) The Data Provider does not perform Code Access Security Link Demands. Two
     instance of Link Demand were removed from the TdQueryBand and TdException
     classes [NET-3962].

  b) A new Statistics schema collection has been added to provide information
     on data distribution within tables [NET-91].

  c) It is no longer required to create mapping tables for Views and ViewColumns
     in order to use the Entity Framework Store Schemas. However better 
     performance will be attained if the mappings are created [NET-3802].

Version 16.0:
  a) The Data Provider supports TDNEGO authentication mechanism in addition to
     LDAP, TD2 and SPNEGO. [NET-3246]

  b) The Data Provider supports the following ADO.NET 4.5 Asynchronous methods:
         1- TdDataReader.ReadAsync
         2- TdDataReader.NextResultAsync
         3- TdDataReader.GetFieldValueAsync<T>  [NET-3484 / NET-3655]
         4- TdCommand.ExecuteNonQueryAsync
         5- TdCommand.ExecuteReaderAsync
         6- TdCommand.ExecuteScalarAsync        [NET-3483]
         7- TdConnection.OpenAsync              [NET-3385]

  c) The Data Provider supports the following ADO.NET 4.5 Streaming methods:
         1- TdDataReader.GetFieldValue<T>
         2- TdDataReader.GetStream    
         3- TdDataReader.GetTextReader
         4- TdDataReader.GetXmlReader           [NET-2481]

  d) TdDataReader.GetSchemaTable() method has been enhanced by adding a new 
     DataColumn for ServerType. This field is populated only when the column is
     a DATASET data type. The data for a Dataset column will be transferred 
     to/from Teradata as a BLOB or VarByte value depending on the transform 
     that is currently in effect for the user. [NET-3480]

  e) The Schema Collections have been enhanced as follows [NET-3707]:

     1- Support has been added for the new DATASET STORAGE FORMAT AVRO data type
        and for the new AVRO SCHEMA object type.

     2- Support has been added for the new INLINE LENGTH attribute that can be 
        applied to columns and parameters of the following data types:
            JSON, XML, ST_GEOMETRY and DATASET.

     3- The Columns collection has been enhanced to return the following 
        additional columns: STORAGE_FORMAT, INLINE_LENGTH and DATASET_SCHEMA

     4- The ProcedureParameters and UserDefinedFunctionParameters collections
        have been enhanced to return the following additional columns:
            STORAGE_FORMAT and INLINE_LENGTH

  f) Support the new 1MB response rows and 16MB messages in Teradata Database 
     16.0. The Data Provider now supports an internal buffer size of up to 16MB.
     This value can be set through the TdConnection.ResponseBufferSize property.
     [NET-3482]


1.2 Problems Fixed
==================

Version 16.20.8:
  a) The Data Provider opens a Connection when System.Environment.UserName
     returns a NULL. The Data Provider no longer throws an NullReference 
     exception when the runtime-environment (e.g. Docker) returns NULL for 
     UserName or UserDomainName. [NET-4679]

Version 16.20.7:
  a) The Data Provider successfully returns a Connection from a Pool when 
     TdConnection.OpenAsync method is used to retrieve an idle Connection. The 
     Data Provider no longer throws "[100002] Cannot create connection within 
     the time specified". [NET-4645] 

Version 16.20.6: 
  a) The Data Provider successfully connects to the Teradata Database when 1) the 
     User's password has expired, 2) the User has specified a new Password
     in the Connection-String and 3) the Teradata Database Network policy is to 
     to always encrypt the Data send to and from the Data Provider. 
     The Data Provider no longer throws a NullReferenceException. [NET-4620] 

Version 16.20.5:
  a) The Data Provider will throw an OperationCancelledException when an
     ExecuteXXXAsync() operation is cancelled using a CancellationTokenSource.
     This will set the Task state to Cancelled. (i.e. Task.IsCancelled will be
     true.) If you call Wait() or WaitAll() on the task it will throw an
     AggregateException that wraps a TaskCancelledException. If you await the 
     task it will throw an OperationCanceledException. [NET-4556]

  b) The Data Provider threw a System.IndexOutOfRangeException when executing
     a SQL query (using TdCommand or TdDataAdapter) with the following 
     characteristics:

         1) one or more Smart-LOB parameters and
         2) the Redrive feature was enabled on the Teradata Database.
     
     This issue is very rare and it has been fixed. [NET-4520]

  c) The Data Provider threw a TdException (error code 5526) from 
     TdCommand.ExecuteCreateProcedure method when the stored-procedure 
     referenced an Extended Object Name (e.g. Column-Name greater than 
     30 characters). This issue has been fixed. [NET-4559]

  d) The Data Provider threw a NullReferenceException when the application 
     invoked TdConnection.Close() method to cancel the reconnection 
     after a network failure. This issue is very rare and it has 
     been fixed. [NET-4560]


Version 16.20.4:
  a) The Data Provider no longer returns an Aborted Connection 
     (see Teradata Database error code 3134) to the pool. This will prevent 
     very rare cases of error codes 3902 (Session is not logged on) or 
     115009 (Message truncation error) when the application retrieved an 
     aborted-connection from a pool. [NET-4331]
     Note: When an administrator aborts an idle pooled connection and 
           the Teradata Database is slow at terminating the socket-connection,
           then the Data Provider might still throw 3902 or 115009.

Version 16.20.3:
  a) The Data Provider no longer throws a NullReferenceException when the 
     Data Provider is published/deployed with the application 
     (private deployment). [NET-4441]
     
     The Data Provider 16.20.2 throws "The type initializer for 
     'Teradata.Client.Provider.UtlPerformanceCounters' threw an exception. 
     ---> System.NullReferenceException: Object reference not set to an 
     instance of an object."


Version 16.20.2:
  a) When inserting a CLOB using a session character set of UTF8 or UTF16 a
     surrogate pair might span 2 messages. In this case each of the surrogate
     pair characters used to be replaced by the 'invalid character'.
     The correct surrogate pair will now be sent. [NET-4201]


Version 16.20.1:
  a) Previously, if an application re-used TdParameter objects associated with 
     a TdCommand in some scenarios while connected to Teradata Database 15.10, 
     the Data Provider may throw TdException 100056: Total size of parameters 
     is greater than the max Data parcel size. This issue has been fixed. 
     [NET-4135]

  b) The Data Provider will no longer attempt to use the DBC.ColumnsJQV view,
     when fetching the Columns Schema Collection, if the Teradata Database QVCI
     feature is disabled. [NET-4210]

  c) Large batch sizes will no longer return Database Error 9993 when using
     the Update() method of a TdDataAdapter to insert, or update, rows on a
     Teradata 16.10 or 16.20 system. [NET-4254]


Version 16.20:
  a) The Data Provider will now correctly report the situation where the Offset 
     specified for an input parameter is larger than the input data value 
     length. [NET-4086]

  b) The Data Provider will now correctly handle an input parameter defined as 
     TdType.VarChar when the parameter value is a TextReader. [NET-4088]

  c) The Data Provider will now correctly ignore the TdParameter.Size value when
     retrieving the value for an output parameter defined as TdType.TdXml. 
     [NET-4119]

  d) If TdParameter.Size is greater than 64000, the Data Provider sends the 
     parameter as a Clob. If the TdParameter.DbType property is set to 
     AnsiString, then TdParameter.Size property refers to the maximum size, in 
     bytes, of the data within the column, matching the DbParameter.Size 
     specification.
     
     Previously, the Data Provider treated the TdParameter.Size of a Clob as 
     the maximum size, in characters, of the data within the column, which did
     not match the DbParameter.Size specification. Depending on the session 
     character set, the Data Provider could send more data to the Teradata 
     Database than the application requested. [NET-3807]

  e) If the session character set is UTF16, TdParameter.DbType property is set 
     to AnsiString, TdParameter.Size is less than or equal to 64000, and the 
     number of characters in TdParametter.Value is less than the number of bytes
     specified in TdParamter.Size, the Data Provider sent fewer characters than
     the application requested. This issue has been fixed. [NET-4046]

  f) A GUI or ASP.NET application may hang in the TdConnection.OpenAsync method.
     This issue has been fixed. [NET-4024]

Version 16.10.1:
  a) The Data Provider will no longer attempt to use the DBC.ColumnsJQV view,
     when fetching the Columns Schema Collection, if the Teradata Database QVCI
     feature is disabled. [NET-4210]

  b) Large batch sizes will no longer return Database Error 9993 when using
     the Update() method of a TdDataAdapter to insert, or update, rows on a
     Teradata 16.10 or 16.20 system. [NET-4254]

Version 16.10:
  a) The Data Provider now supports TCP Receive Window Auto-Tuning feature 
     [NET-3947].
 
  b) The correct error message will now be displayed if a Timeout occurs
     during a batch update operation [NET-3797].

Version 16.0.0.1:
  a) Silent Install uninstalled the previous version of the Data Provider.
     In other words, Silent Install did not support Side-by-Side Installation 
     of two or more versions of the Data Provider. This issue has been fixed 
     [CLNTINS-7280].
      
Version 16.0:

  a) The Data Provider did not calculate the buffer size correctly which 
     resulted in an additional request being sent to the Teradata Database for
     each Batch. This issue has been fixed. [NET-3592]

  b) If a Period column contained a null value in the first row of a batch the 
     Precision and Scale were set to zero. The result was that any fractional 
     seconds were truncated. This issue has been fixed. [NET-3417]


       
2.0 Teradata Product Dependencies
=================================
The Data Provider communicates directly to the Teradata Database. There are no 
dependencies on any of the Teradata Tools and Utilities products.


3.0 Operating Environments
===========================
The following operating environments are supported:

Client Operating Systems:
    Microsoft Windows 7               x86 & x64 Editions
    Microsoft Windows 8               x86 & x64 Editions
    Microsoft Windows 8.1             x86 & x64 Editions
    Microsoft Windows 10              x86 & x64 Editions
    Microsoft Windows Server 2008     x86 & x64 Editions
    Microsoft Windows Server 2008 R2
    Microsoft Windows Server 2012 
    Microsoft Windows Server 2012 R2 
    Microsoft Windows Server 2016
    Microsoft Windows Server 2019


Client Operating Environment:
    Microsoft .NET Framework 4.5.2    x86 & x64 (CLR 4.0)
    Microsoft .NET Framework 4.6.x    x86 & x64 (CLR 4.0)
    Microsoft .NET Framework 4.7.x    x86 & x64 (CLR 4.0)
    Microsoft .NET Framework 4.8.x    x86 & x64 (CLR 4.0)
    Microsoft .NET Core 2.1           Windows, Linux, macOS
    Microsoft .NET Core 2.2           Windows, Linux, macOS
    Microsoft .NET Core 3.1           Windows, Linux, macOS
                    
Microsoft Visual Studio:
    2015 Professional Edition or higher edition
    2017 Professional Edition or higher edition
    2019 Professional Edition or higher edition

The .NET Data Provider for Teradata and .NET Entity Provider for Teradata 
support the following Teradata Database Releases:
    14.10
    15.0
    15.10
    16.10
    16.20
        
The .NET Entity Provider for Teradata supports ADO.NET Entity Framework 3.5 SP1
features. It does not support ADO.NET Entity Framework 4.0 (or higher) features.
The .NET Entity Provider for Teradata is not compatible with 
Entity Framework 6.x or Entity Framework Core. 

The .NET Data Provider for Teradata architecture is MSIL. Therefore a 
single installation supports 32-bit and 64-bit applications.


4.0 Installation
================
For step-by-step Installation Instructions refer to "Teradata Tools and 
Utilities Installation Guide for Microsoft Windows" document on 
https://docs.teradata.com. 


4.1 Installation Procedure
==========================
The .NET Data Provider for Teradata is dependent on the Microsoft .NET Framework
version 4.5.2 which can be downloaded from: 
    https://www.microsoft.com/net/download

Alternatively you may use the .NET Core implementation of the provider. This is
dependent of Microsoft .NET Core version 2.0 which can be downloaded from:
    https://www.microsoft.com/net/download

*** RECOMMENDATION: Uninstall previous version(s) of the .NET Data Provider
for Teradata if there is no application dependency.

In the following instructions, the "installation executable" refers to the 
the file:

     tdnetdp__windows_indep.16.20.09.00.exe

NOTE:
      The name of the installation package obtained when performing 
      a Network Install from the TTU media is setup.exe.   


Any reference to the .NET Data Provider for Teradata or the Teradata Provider
refers to

     .NET Data Provider for Teradata 16.20.09.00

Follow this procedure to install the Data Provider:

   1- Run installation executable.   (See NOTE below.)
   
   2- In the Choose Setup Language dialog, select the Language you want to use
      and click OK.
      
   3- When the Welcome screen appears, click Next>.
   
   4- In the Choose Destination folder dialog, choose a destination folder 
      and click Next>.

   5- In the Setup Type dialog, choose either complete or custom setup type
      and click Next>. Choosing 'complete' setup type would select all the 
      features for installation whereas choosing 'custom' setup type allows 
      the user to select features for installation.

   6- In the Custom setup dialog, click Next> to accept the defaults. 
      The "Integrate the Data Provider with Microsoft Visual Studio(s)" feature 
      is visible when any of the Microsoft Visual Studio versions supported by
      this version of the .NET Data Provider for Teradata is installed. 
      
   7- In the Ready to install dialog, click on 'Install' to start the package 
      installation. If Visual studio features are selected then this step will 
      take several minutes to complete.      
 
   8- In the InstallShield Wizard Complete screen, click Finish.   
   

4.2 Installation options
=========================
Interactive install:
   Double click installation executable to start the installation or Run 
   the installation executable from command prompt and go through the dialog 
   sequence.


Silent install:
  From command prompt:   
    To install the full package specify the installation executable with the
    following default options:

       /s /v"/qn"   

    To install to a user defined location run specify the following options:

       /s /v"/qn INSTALLDIR=c:\test"

    To install the package without Visual Studio features use the options:

       /s /v"/qn VSFEATURE=0"    

NOTE:
   The self-extracting installation package, downloaded from the web
   (downloads.teradata.com), is not compatible with the SYSTEM account.
   The SETUP.EXE installation package, available on the Teradata Tools and
   Utilities media, must be used to deploy the Data Provider with the SYSTEM
   account.


Installation through SMS

   Extract the MSI file from the installation executable by running the 
   following command at command prompt
   
        "installation executable" /a

   Using the extracted '.NET Data Provider for Teradata.msi' file the package 
   can be pushed from SMS server to SMS clients using the following commands:

   To install full package with defaults use
   msiexec /i ".NET Data Provider for Teradata.msi" /qn        

   To install to a user defined location use
   msiexec /i ".NET Data Provider for Teradata.msi" INSTALLDIR=c:\test /qn

   To install the package without Visual Studio features use
   msiexec /i ".NET Data Provider for Teradata.msi" VSFEATURE=0 /qn


5.0 Uninstall
=============
Follow this procedure to uninstall the .NET Data Provider for Teradata: 

  1- Open "Control Panel".

  2- Start "Add or Remove Programs".

  3- Find ".NET Data Provider for Teradata" and click Remove.
     This will uninstall the package.


6.0 Restrictions and Known Issues
=================================
a) The .NET Data Provider for Teradata requires Full-Trust permission-set to 
   run. 
