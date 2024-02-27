//Copyright 2012-2019 Keysight Technologies
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
using OpenTap;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Virinco.Newtonsoft.Json;
using Virinco.WATS.Interface;
using Virinco.WATS.Interface.Models;
using Virinco.WATS.Schemas.WRML;
using static Virinco.WATS.Interface.ReportCenter.WATSReport;

namespace OpenTap.Plugins.PluginDevelopment
{
    [Display("WATSListener", Group: "WATS", Description: "Sends result directly to WATS")]
    public class WatsResultListener : ResultListener
    {
        private TDM api;
        private UUTReport currentWatsUUT;
        private ResultTable currentKeysightResultTable;

        private Dictionary<Guid, TestStepRun> testStepRunLookup = new Dictionary<Guid, TestStepRun>();

        // Create a stack to hold the current sequence hierarchy.
        private Stack<SequenceCall> sequenceStack = new Stack<SequenceCall>();

        // We use these lists to keep track on if a teststep is a flow control step or a test step
        // Flow control steps in OpenTAP will be added as sequences in WATS
        List<Guid> publishedIds = new List<Guid>(); // PublishedIds are the test steps that has used the publish method
        List<Guid> AllIds = new List<Guid>(); // AllIds are all the test steps that has been run
        List<Guid> flowControlIds = new List<Guid>(); // FlowControlIds are the test steps that we have added to wats as sequences

        /////////////////
        // DATA MODELS //
        /////////////////

        private string numericLimitTestsConfigJsonString;
        private Dictionary<string, NumericLimitTestSettings> numericLimitTestSettingsDict = new Dictionary<string, NumericLimitTestSettings>();

        private string stringValueTestsConfigJsonString;
        private Dictionary<string, StringValueTestSettings> stringValueTestSettingsDict = new Dictionary<string, StringValueTestSettings>();

        private string graphTestsConfigJsonString;
        private Dictionary<string, GraphTestSettings> graphTestSettingsDict = new Dictionary<string, GraphTestSettings>();

        private string passFailTestsConfigJsonString;
        private Dictionary<string, PassFailTestSettings> passFailTestSettingsDict = new Dictionary<string, PassFailTestSettings>();

        private string uutFieldsConfigJsonString;
        private Dictionary<string, UUTFieldSettings> uutFieldsSettingsDict = new Dictionary<string, UUTFieldSettings>();

        public class UUTFieldSettings
        {
            public List<string> tests { get; set; }
            public int valueColumn { get; set; }
            public string fieldMapping { get; set; }
            public override string ToString()
            {
                return "ValueColumn: " + valueColumn;
            }
        }

        public class NumericLimitTestSettings
        {
            public List<string> tests { get; set; }
            public int measurmentColumn { get; set; }
            public string compOperatorType { get; set; }
            public int? lowerLimitColumn { get; set; }
            public int? upperLimitColumn { get; set; }
            public int? unitColumn { get; set; }

            public override string ToString()
            {
                return "MeasurmentColumn: " + measurmentColumn + ", compOperatorType: " + compOperatorType + ", LowerLimitColumn: " + lowerLimitColumn + ", upperLimitColumn: " + upperLimitColumn + ", UnitColumn: " + unitColumn;
            }
        }

        public class StringValueTestSettings
        {
            public List<string> tests { get; set; }
            public int valueColumn { get; set; }

            public int? limitColumn { get; set; }
            public override string ToString()
            {
                return "ValueColumn: " + valueColumn + ", limitColumn: " + limitColumn;
            }

        }

        public class GraphTestSettings
        {
            public List<string> tests { get; set; }
            public int xAxisColumn { get; set; }

            public int yAxisColumn { get; set; }

            public override string ToString()
            {
                return "XAxisColumn: " + xAxisColumn + ", YValueColumn: " + yAxisColumn;
            }

        }

        public class PassFailTestSettings
        {
            public List<string> tests { get; set; }
        }

        //////////////////////////
        // SETTING INPUT FIELDS //
        //////////////////////////

        [Display("UUTFields config\n\n Values that can not be null:\n- valueColumn\n-  fieldMapping", Description: "Configure what column are what values for pass/fail tests in Wats", Group: "Config for converting teststeps to WATS", Order: 0)]
        [Layout(LayoutMode.Normal, 10, 10)]
        public string UUTFieldsConfig
        {
            get { return uutFieldsConfigJsonString; }
            set
            {
                try
                {
                    Log.Info("Reading stored JSON for UUT field config");
                    uutFieldsConfigJsonString = value;
                    uutFieldsSettingsDict = new Dictionary<string, UUTFieldSettings>();
                    Log.Info("Attempting to deserialize JSON...");
                    var formats = JsonConvert.DeserializeObject<Dictionary<string, UUTFieldSettings>>(uutFieldsConfigJsonString);
                    Log.Info("JSON deserialized successfully.");
                    Log.Info("Populating settings dictionary...");
                    foreach (var format in formats)
                    {
                        foreach (var test in format.Value.tests)
                        {
                            var setting = format.Value;
                            setting.tests = null; // We don't need the tests inside the setting anymore.

                            // Check if the test step already is in the config settings.
                            if (uutFieldsSettingsDict.ContainsKey(test.Trim()))
                            {
                                Log.Warning("Test step '{0}' is added to mulitple config in settings. The Result listener will use the format that appears at the bottom of the json settings.", test.Trim());
                            }

                            uutFieldsSettingsDict[test.Trim()] = setting;
                        }
                    }
                    Log.Info("Settings dictionary populated successfully.");
                }
                catch (JsonException ex)
                {
                    // This will catch any JSON related errors such as invalid JSON format.
                    Log.Error($"Error occurred while processing JSON: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
                catch (Exception ex)
                {
                    // This will catch any general errors not related to JSON.
                    Log.Error($"An unexpected error occurred: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
            }
        }

        [Display("Numeric limit tests config\n\n Values that can not be null:\n- measurmentColumn\n-  compOperatorType", Description: "Configure what column are what values for numeric limit tests in Wats", Group: "Config for converting teststeps to WATS", Order: 1)]
        [Layout(LayoutMode.Normal, 10, 10)]
        public string NumericLimitTestsConfig
        {
            get { return numericLimitTestsConfigJsonString; }
            set
            {
                try
                {
                    Log.Info("Reading stored JSON for numeric limit tests config");
                    numericLimitTestsConfigJsonString = value;
                    numericLimitTestSettingsDict = new Dictionary<string, NumericLimitTestSettings>();
                    Log.Info("Attempting to deserialize JSON...");
                    var formats = JsonConvert.DeserializeObject<Dictionary<string, NumericLimitTestSettings>>(numericLimitTestsConfigJsonString);
                    Log.Info("JSON deserialized successfully.");
                    Log.Info("Populating settings dictionary...");
                    foreach (var format in formats)
                    {
                        foreach (var test in format.Value.tests)
                        {
                            var setting = format.Value;
                            setting.tests = null; // We don't need the tests inside the setting anymore.

                            // Check if the test step already is in the config settings.
                            if (numericLimitTestSettingsDict.ContainsKey(test.Trim()))
                            {
                                Log.Warning("Test step '{0}' is added to mulitple config in settings. The Result listener will use the format that appears at the bottom of the json settings.", test.Trim());
                            }
                            numericLimitTestSettingsDict[test.Trim()] = setting;
                        }
                    }
                    Log.Info("Settings dictionary populated successfully.");
                }
                catch (JsonException ex)
                {
                    // This will catch any JSON related errors such as invalid JSON format.
                    Log.Info($"Error occurred while processing JSON: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
                catch (Exception ex)
                {
                    // This will catch any general errors not related to JSON.
                    Log.Info($"An unexpected error occurred: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
            }
        }

        [Display("String value tests config\n\n Values that can not be null:\n- valueColumn", Description: "Configure what column are what values for string value tests in Wats", Group: "Config for converting teststeps to WATS", Order: 2)]
        [Layout(LayoutMode.Normal, 10, 10)]
        public string StringValueTestsConfig
        {
            get { return stringValueTestsConfigJsonString; }
            set
            {
                try
                {
                    Log.Info("Reading stored JSON for string value tests config");
                    stringValueTestsConfigJsonString = value;
                    stringValueTestSettingsDict = new Dictionary<string, StringValueTestSettings>();
                    Log.Info("Attempting to deserialize JSON...");
                    var formats = JsonConvert.DeserializeObject<Dictionary<string, StringValueTestSettings>>(stringValueTestsConfigJsonString);
                    Log.Info("JSON deserialized successfully.");
                    Log.Info("Populating settings dictionary...");
                    foreach (var format in formats)
                    {
                        foreach (var test in format.Value.tests)
                        {
                            var setting = format.Value;
                            setting.tests = null; // We don't need the tests inside the setting anymore.

                            // Check if the test step already is in the config settings.
                            if (stringValueTestSettingsDict.ContainsKey(test.Trim()))
                            {
                                Log.Warning("Test step '{0}' is added to mulitple config in settings. The Result listener will use the format that appears at the bottom of the json settings.", test.Trim());
                            }

                            stringValueTestSettingsDict[test.Trim()] = setting;
                        }
                    }
                    Log.Info("Settings dictionary populated successfully.");
                }
                catch (JsonException ex)
                {
                    // This will catch any JSON related errors such as invalid JSON format.
                    Log.Error($"Error occurred while processing JSON: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
                catch (Exception ex)
                {
                    // This will catch any general errors not related to JSON.
                    Log.Error($"An unexpected error occurred: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
            }
        }

        [Display("Graph tests config\n\n Values that can not be null:\n- xAxisColumn\n- yAxisColumn", Description: "Configure what column are what values for graph tests in Wats", Group: "Config for converting teststeps to WATS", Order: 3)]
        [Layout(LayoutMode.Normal, 10, 10)]
        public string GraphTestsConfig
        {
            get { return graphTestsConfigJsonString; }
            set
            {
                try
                {
                    Log.Info("Reading stored JSON for graph tests config");
                    graphTestsConfigJsonString = value;
                    graphTestSettingsDict = new Dictionary<string, GraphTestSettings>();
                    Log.Info("Attempting to deserialize JSON...");
                    var formats = JsonConvert.DeserializeObject<Dictionary<string, GraphTestSettings>>(graphTestsConfigJsonString);
                    Log.Info("JSON deserialized successfully.");
                    Log.Info("Populating settings dictionary...");
                    foreach (var format in formats)
                    {
                        foreach (var test in format.Value.tests)
                        {
                            var setting = format.Value;
                            setting.tests = null; // We don't need the tests inside the setting anymore.

                            // Check if the test step already is in the config settings.
                            if (graphTestSettingsDict.ContainsKey(test.Trim()))
                            {
                                Log.Warning("Test step '{0}' is added to mulitple config in settings. The Result listener will use the format that appears at the bottom of the json settings.", test.Trim());
                            }

                            graphTestSettingsDict[test.Trim()] = setting;
                        }
                    }
                    Log.Info("Settings dictionary populated successfully.");
                }
                catch (JsonException ex)
                {
                    // This will catch any JSON related errors such as invalid JSON format.
                    Log.Error($"Error occurred while processing JSON: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
                catch (Exception ex)
                {
                    // This will catch any general errors not related to JSON.
                    Log.Error($"An unexpected error occurred: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
            }
        }

        [Display("Pass/fail tests config\n\n Has no column configs", Description: "Configure what column are what values for pass/fail tests in Wats", Group: "Config for converting teststeps to WATS", Order: 4)]
        [Layout(LayoutMode.Normal, 10, 10)]
        public string PassFailTestsConfig
        {

            get { return passFailTestsConfigJsonString; }
            set
            {
                try
                {
                    Log.Info("Reading stored JSON for graph tests config");
                    passFailTestsConfigJsonString = value;
                    passFailTestSettingsDict = new Dictionary<string, PassFailTestSettings>();
                    Log.Info("Attempting to deserialize JSON...");
                    var formats = JsonConvert.DeserializeObject<Dictionary<string, PassFailTestSettings>>(passFailTestsConfigJsonString);
                    Log.Info("JSON deserialized successfully.");
                    Log.Info("Populating settings dictionary...");
                    foreach (var format in formats)
                    {
                        foreach (var test in format.Value.tests)
                        {
                            var setting = format.Value;
                            setting.tests = null; // We don't need the tests inside the setting anymore.

                            // Check if the test step already is in the config settings.
                            if (passFailTestSettingsDict.ContainsKey(test.Trim()))
                            {
                                Log.Warning("Test step '{0}' is added to mulitple config in settings. The Result listener will use the format that appears at the bottom of the json settings.", test.Trim());
                            }

                            passFailTestSettingsDict[test.Trim()] = setting;
                        }
                    }
                    Log.Info("Settings dictionary populated successfully.");
                }
                catch (JsonException ex)
                {
                    // This will catch any JSON related errors such as invalid JSON format.
                    Log.Error($"Error occurred while processing JSON: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
                catch (Exception ex)
                {
                    // This will catch any general errors not related to JSON.
                    Log.Error($"An unexpected error occurred: {ex.Message}");
                    Log.Debug($"Full exception details: {ex}");
                }
            }
        }

        ///////////////////////
        // LIFESYCLE METHODS //
        ///////////////////////
        public WatsResultListener()
        {
            api = new TDM();
            api.SetupAPI(null, "WATS", "OpenTap Plugin", true);
            api.InitializeAPI(true);

            // These hardcoded json strings act as the default config settings for the plugin when the dll is loaded for the first time
            // OpenTAP saves theses strings into 'C:\Program Files\OpenTAP\Settings\Results.xml' 
            // If the dll is unloaded and then reloaded again, for instance when you close and reopen the program, it will follow the operations below
            // First the constructor will be called and set the NumericLimitTestsConfig to the default json
            // Then the setters will be called and set the NumericLimitTestsConfig to the saved json from Results.xml
            NumericLimitTestsConfig = "{\r\n   \"format1\":{\r\n      \"tests\":[\r\n         \"NumericalLimitDummyTest1\"\r\n      ],\r\n      \"measurmentColumn\":0,\r\n      \"compOperatorType\":\"EQ\",\r\n      \"lowerLimitColumn\":2,\r\n      \"upperLimitColumn\":null,\r\n      \"unitColumn\":null\r\n   },\r\n   \"format2\":{\r\n      \"tests\":[\r\n         \"NumericalLimitDummyTest2\"\r\n      ],\r\n      \"measurmentColumn\":0,\r\n      \"compOperatorType\":\"GELE\",\r\n      \"lowerLimitColumn\":2,\r\n      \"upperLimitColumn\":3,\r\n      \"unitColumn\":null\r\n   }\r\n}";
            StringValueTestsConfig = "{\r\n   \"format1\":{\r\n      \"tests\":[\r\n         \"StringValueDummyTest\"\r\n      ],\r\n      \"valueColumn\":0,\r\n      \"limitColumn\":1\r\n   }\r\n}";
            GraphTestsConfig = "{\r\n    \"format1\":{\r\n        \"tests\":[\r\n            \"graph\"\r\n        ],\r\n        \"xAxisColumn\":0,\r\n        \"yAxisColumn\":1\r\n    }\r\n}";
            PassFailTestsConfig = "{\r\n    \"format1\":{\r\n        \"tests\":[\r\n            \"passfail\"\r\n        ]\r\n    }\r\n}";
            UUTFieldsConfig = "{\r\n   \"format1\":{\r\n      \"tests\":[\r\n         \"SN\"\r\n      ],\r\n      \"fieldMapping\":\"SerialNumber\",\r\n      \"valueColumn\":0\r\n   },\r\n   \"format2\":{\r\n      \"tests\":[\r\n         \"PN\"\r\n      ],\r\n      \"fieldMapping\":\"PartNumber\",\r\n      \"valueColumn\":0\r\n   }\r\n}";
        }

        public override void Open()
        {
            base.Open();
        }

        public override void Close()
        {
            base.Close();
        }

        public override void OnTestPlanRunStart(TestPlanRun planRun)
        {
            Log.Info("Creating a default Wats UUTReport");
            currentWatsUUT = api.CreateUUTReport("", "", "", "", "10", "", "");
            currentKeysightResultTable = null;
            sequenceStack.Push(currentWatsUUT.GetRootSequenceCall());
            Log.Info("Setting UUTReport StartDateTime: '{0}'", planRun.StartTime);
            currentWatsUUT.StartDateTime = planRun.StartTime;
        }

        public override void OnTestStepRunStart(TestStepRun stepRun)
        {
            // This logging plus the printTable method in the next function makes it easy to see what reults are produced by the teststep
            // By seeing the result in a table format it is easier to configure what columns are what values in the config settings
            Log.Info("");
            Log.Info("Starting step: " + stepRun.TestStepName);

            // We add all the OpenTAP TestStepRun objects to a dictionary so we can acsess it properties another place in the lifecycle of this program
            testStepRunLookup[stepRun.Id] = stepRun;

            // Every teststep ids that is started will be added to the AllIds list
            AllIds.Add(stepRun.Id);
        }

        public override void OnResultPublished(Guid stepRunId, ResultTable resultTable)
        {
            base.OnResultPublished(stepRunId, resultTable);

            // We set the resultTable to a global variable so we can access it in the OnTestStepRunCompleted function
            // We want to handle the result in the OnTestStepRunCompleted function because that is where we can get the verdict of the teststep
            currentKeysightResultTable = resultTable;

            Log.Info("");
            Log.Info("Result:");
            printTable(resultTable);

            // Only the teststep ids that uses this method will be added to the publishedIds list
            publishedIds.Add(stepRunId);
        }

        public override void OnTestStepRunCompleted(TestStepRun stepRun)
        {

            // This is where the logic is difficult to follow

            // A normal step that produces a result will have the following lifecycle: OnTestStepRunStart -> OnResultPublished -> OnTestStepRunCompleted
            // A flow control step will have the following lifecycle: OnTestStepRunStart -> OnTestStepRunCompleted

            // Here is a example of how it looks when a flowcontrolstep with children steps is run:
            // OnTestStepRunStart: SequenceStep
            //      OnTestStepRunStart: NumericTest1
            //      OnResultPublished: NumericTest1
            //      OnTestStepRunCompleted: NumericTest1
            //      OnTestStepRunStart: NumericTest2
            //      OnResultPublished: NumericTest2
            //      OnTestStepRunCompleted: NumericTest2
            // OnTestStepRunCompleted: SequenceStep

            // So what we do is to check if the stepRunId has not used the published method, and has not been added as a flowcontrol step by us already
            // If this is the case we add the step as a sequence in WATS, push it to the sequenceStack and adds the step id to the flowControlIds list
            foreach (Guid id in AllIds)
            {
                if (!publishedIds.Contains(id) && !flowControlIds.Contains(id))
                {
                    Log.Info("Adding '{0}' as a sequence in WATS.", testStepRunLookup[id].TestStepName);
                    SequenceCall sequenceCall = sequenceStack.Peek().AddSequenceCall(testStepRunLookup[id].TestStepName);
                    sequenceCall.StepTime = testStepRunLookup[id].Duration.TotalSeconds;
                    sequenceStack.Push(sequenceCall);

                    flowControlIds.Add(id);
                }
            }

            // If the stepRunId is in publishedIds it means that it is a normal step that produces a result, we therefore create a WATS teststep and add it to the sequence in the top of the sequenceStack
            if (publishedIds.Contains(stepRun.Id))
            {
                if (addResultsToWatsByConfig(currentKeysightResultTable, stepRun))
                {
                    return;
                }
                else
                {
                    Log.Warning("Did not add {0} to Wats UUTReport. Check if configuration for teststep in ResultListener exists and are correct.", stepRun.TestStepName);
                    return;
                }
                // If THE stepRunId is not in the publishedIds list it means that it is a flowcontrol step
                // When a flowcontrol step calls the OnTestStepRunCompleted-method we know that all the children steps have been run and we can pop the sequenceStack
            }
            else
            {
                sequenceStack.Pop();
            }
        }

        public override void OnTestPlanRunCompleted(TestPlanRun planRun, Stream logStream)
        {   
            Log.Info("Setting UUTReport Status: '{0}'", planRun.Verdict);
            currentWatsUUT.Status = getWatsUUTStatusType(planRun.Verdict);
            
            Log.Info("Setting UUTReport ExecutionTime: '{0}'", planRun.Duration.TotalSeconds);
            currentWatsUUT.ExecutionTime = planRun.Duration.TotalSeconds;
            
            Log.Info("Sending UUTReport to Wats");
            api.Submit(currentWatsUUT);
        }

        //////////////////
        // WATS METHODS //
        //////////////////

        private bool addResultsToWatsByConfig(ResultTable resultTable, TestStepRun stepRun)
        {
            try
            {
                string stepName = stepRun.TestStepName.Trim();
                if (uutFieldsSettingsDict.ContainsKey(stepName))
                {
                    UUTFieldSettings settings = uutFieldsSettingsDict[stepName];
                    if (uutFieldsSettingsDict[stepName].fieldMapping == "SerialNumber")
                    {
                        var SerialNumber = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport SerialNumber: '{0}'", SerialNumber);
                        currentWatsUUT.SerialNumber = SerialNumber;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "PartNumber")
                    {
                        var PartNumber = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport PartNumber: '{0}'", PartNumber);
                        currentWatsUUT.PartNumber = PartNumber;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "Location")
                    {
                        var Location = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport Location: '{0}'", Location);
                        currentWatsUUT.Location = Location;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "OperationType")
                    {
                        var OperationType = api.GetOperationType(resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString());
                        Log.Info("Setting UUTReport OperationType: '{0}'", OperationType);
                        currentWatsUUT.OperationType = OperationType;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "BatchSerialNumber")
                    {
                        var BatchSerialNumber = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport BatchSerialNumber: '{0}'", BatchSerialNumber);
                        currentWatsUUT.BatchSerialNumber = BatchSerialNumber;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "SequenceName")
                    {
                        var SequenceName = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport SequenceName: '{0}'", SequenceName);
                        currentWatsUUT.SequenceName = SequenceName;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "SequenceVersion")
                    {
                        var SequenceVersion = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport SequenceVersion: '{0}'", SequenceVersion);
                        currentWatsUUT.SequenceVersion = SequenceVersion;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "PartRevisionNumber")
                    {
                        var PartRevisionNumber = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport PartRevisionNumber: '{0}'", PartRevisionNumber);
                        currentWatsUUT.PartRevisionNumber = PartRevisionNumber;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "StationName")
                    {
                        var StationName = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport StationName: '{0}'", StationName);
                        currentWatsUUT.StationName = StationName;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "Purpose")
                    {
                        var Purpose = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport Purpose: '{0}'", Purpose);
                        currentWatsUUT.Purpose = Purpose;
                    }
                    else if (uutFieldsSettingsDict[stepName].fieldMapping == "Operator")
                    {
                        var Operator = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();
                        Log.Info("Setting UUTReport Operator: '{0}'", Operator);
                        currentWatsUUT.Operator = Operator;
                    }
                    else
                    {
                        Log.Info("FieldMapping '{0}' does not exist in Wats, choose between (\"SerialNumber\", \"PartNumber\", \"Location\", \"OperationType\", \"BatchSerialNumber\", \"SequenceName\", \"SequenceVersion\", \"PartRevisionNumber\", \"StationName\", \"Purpose\", \"Operator\")", uutFieldsSettingsDict[stepName].fieldMapping);
                    }
                }
                else if (numericLimitTestSettingsDict.ContainsKey(stepName))
                {
                    Log.Info("Found settings for '{0}' in Numeric Limit Test Settings dictionary.", stepName);
                    NumericLimitTestSettings settings = numericLimitTestSettingsDict[stepName];

                    string unit;
                    double measurment = double.Parse(resultTable.Columns[(int)settings.measurmentColumn].Data.GetValue(0).ToString());
                    CompOperatorType? comp = getWatsCompOperatorType(settings.compOperatorType);

                    if (settings.unitColumn != null)
                    {
                        unit = resultTable.Columns[(int)settings.unitColumn].Data.GetValue(0).ToString();
                    }
                    else
                    {
                        unit = "";
                    }

                    if (stepRun.Verdict != Verdict.NotSet)
                    {
                        NumericLimitStep currentWatsStep = sequenceStack.Peek().AddNumericLimitStep(stepName);
                        if (comp == null || comp == CompOperatorType.LOG || comp == CompOperatorType.CASESENSIT || comp == CompOperatorType.IGNORECASE)
                        {
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, unit, getWatsStepStatusType(stepRun.Verdict));
                        } 
                        else if (comp == CompOperatorType.EQ)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.EQ, lowerLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        } 
                        else if (comp == CompOperatorType.GT)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GT, lowerLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.GE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GE, lowerLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LT)
                        {
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LT, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LE)
                        {
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LE, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.GELE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GELE, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.GTLT)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GTLT, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.NE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.NE, lowerLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.GELT)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GELT, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.GTLE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.GTLE, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LTGT)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LTGT, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LEGE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LEGE, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LEGT)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LEGT, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else if (comp == CompOperatorType.LTGE)
                        {
                            double lowerLimit = double.Parse(resultTable.Columns[(int)settings.lowerLimitColumn].Data.GetValue(0).ToString());
                            double upperLimit = double.Parse(resultTable.Columns[(int)settings.upperLimitColumn].Data.GetValue(0).ToString());
                            NumericLimitTest currentWatsTest = currentWatsStep.AddTest(measurment, CompOperatorType.LTGE, lowerLimit, upperLimit, unit, getWatsStepStatusType(stepRun.Verdict));
                        }
                        currentWatsStep.StepTime = stepRun.Duration.TotalSeconds;
                    }
                    else
                    {

                        Log.Info("The settings for '{0}' did not meet the conditions to add a numeric test.", stepName);
                    }
                    


                }
                else if (stringValueTestSettingsDict.ContainsKey(stepName))
                {
                    Log.Info("Found settings for '{0}' in String Value Test Settings dictionary.", stepName);

                    StringValueTestSettings settings = stringValueTestSettingsDict[stepName];

                    StringValueStep currentWatsStep = sequenceStack.Peek().AddStringValueStep(stepName);

                    string value = resultTable.Columns[settings.valueColumn].Data.GetValue(0).ToString();

                    if (stepRun.Verdict != Verdict.NotSet)
                    {
                        if (settings.limitColumn == null)
                        {
                            StringValueTest currentWatsTest = currentWatsStep.AddTest(value, getWatsStepStatusType(stepRun.Verdict));
                        }
                        else
                        {
                            string limit = resultTable.Columns[(int)settings.limitColumn].Data.GetValue(0).ToString();
                            StringValueTest currentWatsTest = currentWatsStep.AddTest(CompOperatorType.EQ, value, limit, getWatsStepStatusType(stepRun.Verdict));
                        }
                    }  

                    currentWatsStep.StepTime = stepRun.Duration.TotalSeconds;

                    Log.Info("Added '{0}' as a string value test.", stepName);
                }
                else if (passFailTestSettingsDict.ContainsKey(stepName))
                {
                    Log.Info("Found settings for '{0}' in Pass/fail Test Settings dictionary.", stepName);
                    PassFailStep currentWatsStep = sequenceStack.Peek().AddPassFailStep(stepName);
                    PassFailTest currentWatsTest = currentWatsStep.AddTest(getWatsStepStatusType(stepRun.Verdict) == StepStatusType.Passed);
                    currentWatsStep.StepTime = stepRun.Duration.TotalSeconds;
                    Log.Info("Added '{0}' as a pass/fail test.", stepName);
                }
                else if (graphTestSettingsDict.ContainsKey(stepName))
                {
                    GraphTestSettings settings = graphTestSettingsDict[stepName];
                    GenericStep currentWatsStep = sequenceStack.Peek().AddGenericStep(GenericStepTypes.Action, resultTable.Name);
                    Chart chart = currentWatsStep.AddChart(ChartType.Line, resultTable.Name, resultTable.Columns[0].Name, "", resultTable.Columns[1].Name, "");

                    Array xAxisNonGeneric = resultTable.Columns[settings.xAxisColumn].Data;
                    Array yAaxisNonGeneric = resultTable.Columns[settings.yAxisColumn].Data;

                    double[] xAxisGeneric = xAxisNonGeneric.OfType<object>().Select(o => double.Parse(o.ToString())).ToArray();
                    double[] yAaxisGeneric = yAaxisNonGeneric.OfType<object>().Select(o => double.Parse(o.ToString())).ToArray();

                    chart.AddSeries(resultTable.Name, xAxisGeneric, yAaxisGeneric);
                }
                else
                {
                    return false;
                }
                return true;
            }
            catch (IndexOutOfRangeException ex)
            {
                Log.Error($"Table column from config settings does not exist: {ex.Message}");
                Log.Debug($"Full exception details: {ex}");
                return false;
            }
            catch (FormatException ex)
            {
                Log.Error($"Table column from config settings has the wrong type: {ex.Message}");
                Log.Debug($"Full exception details: {ex}");
                return false;
            }
            catch (Exception ex)
            {
                Log.Error($"An unexpected error occurred: {ex.Message}");
                Log.Debug($"Full exception details: {ex}");
                return false;
            }
        }


        //////////////////////
        // HELPER FUNCTIONS //
        //////////////////////
        public static short ParseStringToShort(string input, short defaultValue = 0)
        {
            if (short.TryParse(input, out short result))
            {
                return result;
            }
            else
            {
                return defaultValue;
            }
        }

        private CompOperatorType? getWatsCompOperatorType(string compOperatorTypeString)
        {
            if (Enum.TryParse(compOperatorTypeString, out CompOperatorType compOperatorTypeEnum))
            {
                Log.Info("compOperatorTypeString '{0}' was sucsesfully recognized", compOperatorTypeString);
                return compOperatorTypeEnum;
            }
            else
            {
                Log.Error("compOperatorTypeString '{0}' is not recognized, choose between (\"EQ\", \"NE\", \"GT\", \"LT\", \"GE\", \"LE\", \"GTLT\", \"GELE\", \"GELT\", \"GTLE\", \"LTGT\", \"LEGE\", \"LEGT\", \"LTGE\", \"LOG\", \"CASESENSIT\", \"IGNORECASE\")", compOperatorTypeString);
                return null;
            }
        }

        private StepStatusType getWatsStepStatusType(Verdict verdict)
        {
            switch (verdict)
            {
                case Verdict.Pass:
                    return StepStatusType.Passed;
                case Verdict.Fail:
                    return StepStatusType.Failed;
                case Verdict.Error:
                    return StepStatusType.Error;
                case Verdict.Aborted:
                    return StepStatusType.Terminated;
                case Verdict.Inconclusive:
                    return StepStatusType.Done;
                case Verdict.NotSet:
                    return StepStatusType.Skipped;
                default:
                    return StepStatusType.Passed;
            }
        }

        private UUTStatusType getWatsUUTStatusType(Verdict verdict)
        {
            switch (verdict)
            {
                case Verdict.Pass:
                    return UUTStatusType.Passed;
                case Verdict.Fail:
                    return UUTStatusType.Failed;
                case Verdict.Error:
                    return UUTStatusType.Error;
                case Verdict.Aborted:
                    return UUTStatusType.Terminated;
                default:
                    return UUTStatusType.Passed;
            }
        }

        private void printTable(ResultTable resultTable)
        {
            int numberOfColumns = resultTable.Columns.Length;
            int numberOfRows = resultTable.Rows;
            int columnWidth = 20;

            // Print the upper border of the table
            Log.Info(new string('-', numberOfColumns * columnWidth + numberOfColumns + 1));

            // Print column names
            string line = "|";
            for (int l = 0; l < numberOfColumns; l++)
            {
                string name = resultTable.Columns[l].Name;
                line += name.PadRight(columnWidth) + "|";
            }
            Log.Info(line);

            // Print a separator line
            Log.Info(new string('-', numberOfColumns * columnWidth + numberOfColumns + 1));

            // Print rows
            for (int d = 0; d < numberOfRows; d++)
            {
                line = "|";
                for (int l = 0; l < numberOfColumns; l++)
                {
                    string value = resultTable.Columns[l].Data.GetValue(d) + "";
                    line += value.PadRight(columnWidth) + "|";
                }
                Log.Info(line);
            }

            // Print the lower border of the table
            Log.Info(new string('-', numberOfColumns * columnWidth + numberOfColumns + 1));
            Log.Info("");
        }

    }

}
