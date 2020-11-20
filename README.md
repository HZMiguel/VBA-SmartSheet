# VBA-SmartSheet
Set of functions that runs SmartSheet API through Excel VBA

## Requirements
### 1. VBA JSON Parser
This library will be used to deal with JSON responses from Smartsheet API.
- Download from original creator: [VBA-JSON-PARSER](https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a)

Import module [Module_VBA_JSON_PARSER.bas](https://github.com/HZMiguel/VBA-SmartSheet/blob/main/Module_VBA_JSON_PARSER.bas) into Excel VBA.
### 2. Excel VBA References
The following VBA libraries must be activated. 
> VBA > Tools > References
- Visual Basic for Applications
- Microsoft Excel 16.0 Object Library
- OLE Automation
- Microsoft XML, v3.0
- Microsoft Scripting Runtime
- Microsoft Script Control 1.0
- Microsoft ActiveX Data Objects (Multidimensional) 2.0 
### 3. Dedicated Token from Smartsheet.
Retrieve the user token from:
> Smartsheet > Personal Settings > API Access > Generate Token
