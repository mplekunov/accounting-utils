This program will allow you to find differences between two files. It will create a new (output) file which will contain the specific data that is different.

In order to use this program you have to correctly setup the config.json file.

Config file has the following format

```
[
    {
        "searchFileName": "searchFile",
        "statementFileName": "statementFile",
        "outputFileName": "outputFile"
    }
]
```

The ```searchFileName``` and ```statementFileName``` can be any name of the file (including the full name of the file with excel extension).
The ```outputFileName``` is the name of the excel file that will contain the output information regarding the differences. This file name SHOULD NOT have any extension provided to it.

Program supports processing of multiple file pairs at the same time.

In order to process multiple files, add new blocks of 
``` 
{
  "searchFileName": "searchFile2",
  "statementFileName": "statementFile2",
  "outputFileName": "outputFile2"
}
```
inside of the ```[]``` square brackets like in the following example:

```
[
    {
        "searchFileName": "searchFile",
        "statementFileName": "statementFile",
        "outputFileName": "outputFile"
    },
    {
        "searchFileName": "searchFile2",
        "statementFileName": "statementFile2",
        "outputFileName": "outputFile2"
    }
]
```
Don't forget to put comma ```,``` between blocks of objects. 
