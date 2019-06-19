# Age-of-air-Calc
Advanced analysis of tracer gas measurements to calculate the efficiency of ventilation in rooms.

### Installation and activation of Age-of-air-calc
- Simply download the spreadsheet file and open it in Microsoft Excel. No installation or registration is needed.
- However, this is a Visual Basic macro-enabled spreadhseet. You must activate macros for it to function: 
  - When you open the file for the first time in Excel, you will see a yellow bar at the top of the window, with the message *"PROTECTED VIEW Be careful... [Enable Editing]"*. Click on the 'Enable Editing' button. 
  - Next, depending on the security settings on your installation of Microsoft Excel, a red bar may appear at the top of the window, with the message *"BLOCKED CONTENT Macros in this document have been disabled..."*. This can be solved by moving the file to a directory on your PC that you designate for files that you trust, then open the file in Excel. To designate a 'trusted directory', click on **File > Options > Trust Center > Trust Center Settings > Trusted Locations > Add new location**, then browse to a directory, e.g. C:\TEMP\. Finally check that **Trust Center Settings > Trusted Documents > Disable Trusted Documents**  is not ticked.
  - When macros are properly activated, **Age-of-air-CalcL** shows a small square splash-screen when you open the file. This splash screen shows the licence info, and advises you if an update is available for download from GitHub. Simply press 'Close' button to close the splash-screen. 
  - If still no splash-screen appears, then you might be able to fix it by menu option **File > Options > Trust Center > Trust Center Settings > Macro Settings > Disable all macros with notification**, which is a suitable level of security.

### Output features
- Calculates local air change index (e.g. in the breathing zone), and air change efficiency for the room as a whole, based on age-of-air concepts. These two incides are defined in a number of standards, e.g. Nordtest NT VVS 047 *"Mean age of air"*, and ASHRAE Standard 129 *" Measuring Air Change Effectiveness"*. A good introductory guide to these concepts is the AIVC Guide *"A Guide to Energy-Efficient Ventilation"* (www.aivc.org).
- Also calculates the ventilation rate (volumetric flow rate and air change rate) in the room based on two methods, (i) the measured time-constant of the change of concentration, and (ii) the measured dosing rate and steady-state concentation of tracer gas in the room. Both methods are defined in ISO 12569 *"Determination of specific airflow rate in buildings - Tracer gas dilution method"*, and ASTM E741 *"Standard Test Method for Determining Air Change in a Single Zone by Means of a Tracer Gas Dilution"*.

### Measurement/calculation capabilities
- Both step-up and step-down experiments can be analyzed.
- It can import measurements taken by Brüel & Kjær / Innova / Lumasense (www.lumasenseinc.com) Multigas Monitor 1314/1412/1512 with Multipoint Doser and Sampler 1303/1403. Alternatively, you can also paste in data from other sources.
- Up to 6 sampling points are presently supported. (This could be extended to 12 or more if GitHub users pull-request). You designate one of these points as a background concentration, and another as the exhaust duct concentration. The remaining sampling points can for example be located thoughout the breathing zone.
- The software automatically corrects for background concentration of tracer gas (e.g. outdoor concentration, or recirculation in the duct system). This uses a more advanced/accurate advanced in existing standards and textbooks, accounting for the fact that the background concentration can be changing over time with a phase lag between supply and exhaust.
- It can be used with any useful tracer gas, e.g. SF<sub>6</sub> (The gas filter number can be changed in 'Global settings' on sheet 'Instructions')

### Experimental setup
- It is recommended to use at least two sampling pointss (one in the supply duct before dosing point, and one in the extract duct). Optional additional sampling points may be located in the breathing zone.
- It is recommended to have as rapid logging as possible. This is achieved by limiting the number of sampling points (2 or 3), and using maximum two gas filters (e.g. SF<sub>6</sub> and water vapour for cross-contamination correction).
- If the room has multiple air terminal devices, then both the background (ambient) sampling, and dosing should be located in the combined flow in the main duct, before it splits to the different air terminal devices in the room. It is very important that the dosing gas gets properly mixed upstream of the air terminal devices to the room, so that the air supplied to the room has a uniform concentration. One means of achieving this is to dose the gas via a flow cross (an averaging pitot station with multiple holes) in the duct.
- The the background (ambient) concentration should be measured upstream of the dosing point, so that it does not measure the dosed gas.
- If the room has multiple exhaust terminals, the extract concentration should be measured in the combined flow in the main extract duct, where the flows from the different extract terminals has joined.
- The ventilation rate should be approximately constant thoughout the measurement period. This is a fundamental assumption for the step-up and step-down methods.
- Use the correct dosing rate of tracer gas so that the concentration lies within the optimal linear region of calibration of the Multigas Monitor. Too high concentration of tracer gas may lead to separation effects (due to difference in density of air and tracer gas).

### Experiment tips
- It is sensible to run a step-up immediately followed by a step-down measurement. In this way, the tracer gas is used for two useful measurements instead of just one, at no additional cost. My experience is that the step-down measurement is often more accurate than the step up, especially in cases when it was difficult to achieve perfectly uniform concentration in the supply air during the step-up.
- If the room has other air inflow points (e.g. leakage), then the step-up method is not suitable for analysis. Analyse the step-down only. It is *not* recommended to operate a fan in the room to quickly achieve a fully uniform concentration before the step-down. This is because the fan will disturb the normal flow pattern in the room, especially if is is strongly influenced by bouyancy driven flow, so as not to disturb stratification, with and temperature stratification.
- If the room has other air outflow points (e.g. leakage), then it is still OK to use both step-up and step-down methods. In either method, the room air change efficiency (ε<sub>a</sub>) may be inaccurate, as there is no combined exhaust. One option is to measure the concentration of the combined flows in the duct system (However you should then subtract the time it takes for the air to flow from the exhaust grilles to the sampling point),

### Step-by-step instructions
- **Step 1**: Open the gas monitor software 7620/7650. Open the database (either a Ventilation or Monitoring database, if it isn't already open). Use menu option Setup > Units to check that the tracer gas concentration is output in units mg/m³ and dosing rate to mg
- **Step 2**: Use menu option File > Close to close the database. By the way, this does not exit the software. Then use menu option File > Export and choose to export all the measured channels plus setup info. This exports text files to your hard disk.
- **Step 3**: Locate the exported text files. They will normally be saved in *Program files* > *7620* > *dba* or *dbm* > (name of database). The setup files are named "bkamalse.txt" for ventilation and "bkmmalse" for monitoring. Each channel has an output file "bkairm*.txt" for ventilation or "bkmonm*.txt" for monitoring.
- **Step 4**: Open this Excel workbook. Make sure that macros are activated. A message box should appear when you open the file, showing that the macro is activated.
- **Step 5**: To import the text files into this workbook, click on the "Import" button in the "Data" worksheet. Text file location was found in Step 2.
- **Step 6**: To analyse the measurements, go to worksheet "Analyze". You should adjust the values in blue text in the column entitled "Analysis settings", in order to fine-tune the analysis. By definition, the calculated local air change index (ε<sub>p</sub>) in the exhaust duct should be exactly 100%. One reason for a deviation from 100% may be the presence of air leakages (infiltration and exfiltration) in the room boundaries, so adjust the ventilation flow rate to account for these, and to get ε<sub>p</sub>=100%.

### License & warranty
- Distributed free under the "CC BY-SA 4.0" lisence (https://creativecommons.org/licenses/by-sa/4.0/). BY-SA is Attribute-ShareAlike but does not require source-code disclosure.
- Please acknowledge/attribute use of this software in your report/publication with an appropriate author citation and URL to this site.
- Provided without warranty of any kind.

### Authour and copyright owner
Peter.Schild@OsloMet.no
