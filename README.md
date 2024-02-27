1. Copy the following DLLs to the `Packages` directory located at `C:\Program Files\Keysight\Test Automation\Packages` on Windows. These DLLs can be found in the `1.0.0` release zip folder:
    - `Virinco.WATS.Interface.TDM.dll`
    - `Virinco.WATS.WATS-Core.dll`
    - `Virinco.Newtonsoft.Json.dll`
    - `OpentapResultListener.dll`
    - `OpentapDummySteps.dll` (For demo purposes)

2. Register for a trial user and install the WATS Client. Here’s a link for the WATS Client download as well as the Release Notes, Installation Guide, API Reference, etc…  [support.virinco.com ](https://support.virinco.com/hc/en-us/articles/207491836-Download-the-WATS-Client)

3. Open Keysight Test Automation and add the WATSListener to your test sequence. Run the test sequence and filter the console output to display logs from this listener. The logs will present the output of each test step in a table format. You will see a warning indicating that a step result was not sent to WATS due to incomplete configuration. To transmit data to WATS, configure the result listener settings for each test step in JSON format. Below is how to set up these configurations.

4. In the 1.0.0 release zip folder, you will find "WatsDemoRun.TapPlan". This contains a demo test sequence that demonstrates how mapping is performed.  



# Config format overview

## Key components

- **`<FormatName>`**: A unique label for each format, serving as a name tag to identify and organize different settings for various tests.

- **`tests`**: A list of specific test steps that the format applies to, indicating to the system which tests should use this format.


## String value tests config

    {
        "<FormatName>": {
            "tests": [
                "<TestStepName>",
                "<TestStepName>",
                "<TestStepName>"
            ],
            "valueColumn": <ColumnIndex>,
            "limitColumn": <ColumnIndex>
        },
        "<FormatName>": {
            "tests": [
                "<TestStepName>"
            ],
            "valueColumn": <ColumnIndex>,
            "limitColumn": <ColumnIndex>
        }
    }

- **`<ColumnIndex>`**: An integer that specifies the column index for test values. `valueColum` is required and can not be set to `null`.


## Numerical limit tests config

    {
        "<FormatName>": {
            "tests": [
                "<TestStepName>",
                "<TestStepName>",
                "<TestStepName>"
            ],
                "measurmentColumn": <ColumnIndex>,
                "compOperatorType": "<compOperatorType>",
                "lowerLimitColumn": <ColumnIndex>,
                "upperLimitColumn": <ColumnIndex>,
                "unitColumn": <ColumnIndex>
        },
        "<FormatName>": {
            "tests": [
                "<TestStepName>"
            ],
                "measurmentColumn": <ColumnIndex>,
                "compOperatorType": "<CompOperatorType>",
                "lowerLimitColumn": <ColumnIndex>,
                "upperLimitColumn": <ColumnIndex>,
                "unitColumn": <ColumnIndex>
        }
    }

- **`ColumnIndex`**: An integer that specifies the column index for test values. `measurmentColumn` is required and can not be set to `null`.
   **`<CompOperatorType>`**: A string representing the type of limit test. Possible types include: `EQ`, `NE`, `GT`, `LT`, `GE`, `LE`, `GTLT`, `GELE`, `GELT`, `GTLE`, `LTGT`, `LEGE`, `LEGT`, `LTGE`, `LOG`, `CASESENSIT`, `IGNORECASE`.



## Pass/fail tests config

    {
        "<FormatName>":{
            "tests":[
                "<TestStepName>",
                "<TestStepName>",
                "<TestStepName>"
            ]
        }
    }

- For this configuration, no columns need to be set since only the verdict of the test step is used as the status for the WATS step.

## UUT fields config

    {
        "<FormatName>":{
            "tests":[
                "<TestStepName>"
            ],
            "fieldMapping": "<FieldMapping>",
            "valueColumn": <ColumnIndex>
        },
        "<FormatName>":{
            "tests":[
                "<TestStepName>"
            ],
            "fieldMapping": "<FieldMapping>",
            "valueColumn": <ColumnIndex>
        }
    }

- **`<FieldMapping>`**: A string indicating which WATS UUT field the result should be mapped to. Possible values include: `SerialNumber`, `PartNumber`, `Location`, `OperationType`, `BatchSerialNumber`, `SequenceName`, `SequenceVersion`, `PartRevisionNumber`, `StationName`, `Purpose`, `Operator`.

- **`ColumnIndex`**: An integer that specifies the column index for test values. `valueColumn` is required and can not be set to `null`.


## Graph tests

    {
        "<FormatName>":{
            "tests":[
                "<TestStepName>",
                "<TestStepName>"
            ],
            "xAxisColumn": <ColumnIndex>,
            "yAxisColumn": <ColumnIndex>
        }
    }


- **`ColumnIndex`**: An integer that specifies the column index for test values. `xAxisColumn` and `yAxisColumn` is required and can not be set to `null`.