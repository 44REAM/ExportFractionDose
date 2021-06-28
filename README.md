

This code is for extract daily dose from eclipse varian treatment planning system (Fraction.cs).
Output file are CT image (ct.vtk), Dose image (dose_n.vtk), Structure and fraction information (fraction.txt).
In file (fraction.txt) are organize type of dose fraction per line. For example:

```
9 2 dose_1 dose_3 
16 2 dose_2 dose_3 
3 1 dose_4
```

This is mean there are three type (3 lines) of dose distribution in treatment course.
Three type of dose distribution come from combination of four dose distribution.
First number is number of fraction, second number is number of dose distribution use for combination.
First line mean this dose deliver for 9 fraction using combination of dose from "dose_1" and "dose_3".
Dose, CT and Structure extract from this code may result in different origin and spacing, further preprocessing may require.

Note that this code can be extract only one course per patient. Patient that had multiple course per single treatment
were extracted manualy.

Export3D.cs are copy from https://github.com/VarianAPIs/Varian-Code-Samples/blob/master/Eclipse%20Scripting%20API/plugins/Export3D.cs .
