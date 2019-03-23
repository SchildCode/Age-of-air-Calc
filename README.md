# Age-of-air-Calc
Advanced analysis of tracer gas measurements to calculate the efficiency of ventilation in rooms.

### Features
- Calculates both local air change index (e.g. in the breathing zone) and air change efficiency for the room as a whole, based on age-of-air concepts. These incides are defined in a number of standards, e.g. Nordtest NT VVS 047 "Mean age of air"
- It can import measurements taken by Brüel & Kjær / Innova / Lumasense (www.lumasenseinc.com) multigas monitor 1314/1412/1512 with Multipoint Doser and Sampler 1303/1403, but you can also paste in data from other sources
- Up to 6 sampling points are presently supported. (This could be extended to 12 or more if GitHub users pull-request).
- It can be used with any useful tracer gas, e.g. SF6 (The gas filter number can be changed in 'Global settings' on sheet 'Instructions')
- The software automatically corrects for background concentrations, which can change over time.
- 

### Experimental setup
- It is recommended to sample using at least two channels (one in the supply duct before dosing point, and one in the extract duct), and optional additional sampling points located in the breathing zone.
- If the room has multiple air terminal devices, then both the background (ambient) sampling, and dosing should be located in the combined flow in the main duct, before it splits to the different air terminal devices in the room.
- The the background (ambient) concentration should be measured upstream of the dosing point, so that it does not measure the dosed gas.
- If the room has multiple exhaust terminals, the extract concentration should be measured in the combined flow in the main extract duct, where the flows from the different extract terminals has joined.

### Experimantal method
- It is sensible to run a step-up immediately followed by a step-down measurement. In this way, the tracer gas is used for two useful measurements instead of just one, at no additional cost.
- If the room has other air inflow points (e.g. leakage), then the step-up method is not recommented. Use the step-down method after having mixed the room air with a fan. (If bouyancy driven flow, use fan sparingly, so as not to disturb stratification)
- If the room has other air outflow points (e.g. leakage), then it is still OK to use both step-up and step-down methods. In either method, the room air change efficiency (εa) may be inaccurate, as there is no combined exhaust.

### Instructions
- **Step 1**: Open the gas monitor software 7620/7650. Open the database (either a Ventilation or Monitoring database, if it isn't already open). Use menu option Setup > Units to check that the tracer gas concentration is output in units mg/m³ and dosing rate to mg
- **Step 2**: Use menu option File > Close to close the database. By the way, this does not exit the software. Then use menu option File > Export and choose to export all the measured channels plus setup info. This exports text files to your hard disk.
- **Step 3**: Locate the exported text files. They will normally be saved in Program files > 7620 > dba or dbm > (name of database). The setup files are named "bkamalse.txt" for ventilation and "bkmmalse" for monitoring. Each channel has an output file "bkairm*.txt" for ventilation or "bkmonm*.txt" for monitoring.
- **Step 4**: Open this Excel workbook. Make sure that macros are activated. A message box should appear when you open the file, showing that the macro is activated.
- **Step 5**: To import the text files into this workbook, click on the "Import" button in the "Data" worksheet. Text file location was found in Step 2.
- **Step 6**: To analyse the measurements, go to worksheet "Analyze". You should adjust the values in blue text in the column entitled "Analysis settings", in order to fine-tune the analysis.
