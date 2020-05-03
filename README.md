
# Perfomance testing for Excel calculators

This is a rudimentary collection of scripts which help set up a benchmark or series of benchmarks that can exercise various Python solutions evaluating formula networks found within Excel files.

This isn't a "click-and-wait" situation. There is opportunity to configure benchmarking Excel files to exercise the libraries in various ways. I am very comfortable accepting PRs to formalise particular benchmarks to develop a library of challenges.

The current challenge is a series of heavily nested SUM operations, and volume of them.

Currently the Python solutions tested are;
- PyCel
- Koala2
- xlcalcualtor

The libraries Formulas and Schedula have been considered but are yet to be integrated into the tests.

Each of these solutions have a different focus and so are expected to have differing strengths and weaknesses.

# The flow
The script prepare_test_xlsx.py uses pyopenxl to generate a "template" xlsx file. This template file has prepared formulas but no cells have an initial value.

The structure of the file is such;
- Row 1 are headers
- Row 2 and beyond (to a configurable row count)
  - A2, B2, C2 get set as random numbers (0, 1) before the test begins.
  - D2 is a formula =SUM(A2, B2:C2, 10)
  - Columns E through Z use the SUM function to sum from A through to the column before current eg;
    - E2 =SUM(A2:D2)
    - F2 =SUM(A2:E2)
    - G2 =SUM(A2:F2)
    - ...

The function SUM was chosen as we can exercise the Python libraries with a cell reference, a range and a number in a single step and all known libraries implement the SUM function.

Nesting the functions from columns E through to Z provides a challenge for recursion or otherwise following dependencies of a given cell. It's not possible to get a correct answer in a given column without calculating the preceding columns.

Providing a large number of rows using the above pattern tests file loading techniques and evaluation throughput.

Before being used, the template file has column A, B and C prepared with random numbers (0, 1) for all rows taking part in the test.

Then a given library is set to work loading the file and running the appropriate "evaluate" method on pre-defined cells. The pre-defined cells can be carefully chosen to test calculation throughput, caching or cell dependency management.


# Notes for Koala2

I have only been able to get Koala2 0.0.33 running using Python 2.7 in a virtualenv. Koala2 has released 0.0.35 but I wasn't able to get that working so any test results discussed are using an older version of Koala2. This is also the reason Koala2 isn't in setup.py or requirements.py.


## PyCel

```
python PyCel_individual_SUM.py
Loading Excel_individual_SUM_10000.xlsx...
ExcelCompiler made 0:00:00.563514
addresses made 0:00:00.563514
EVALUATED VALUE Sheet1!D2 12.502811628302023
EVALUATED VALUE Sheet1!E2 14.5727862427451
EVALUATED VALUE Sheet1!F2 29.1455724854902
EVALUATED VALUE Sheet1!G2 58.2911449709804
EVALUATED VALUE Sheet1!H2 116.5822899419608
EVALUATED VALUE Sheet1!I2 233.1645798839216
EVALUATED VALUE Sheet1!J2 466.3291597678432
EVALUATED VALUE Sheet1!K2 932.6583195356864
EVALUATED VALUE Sheet1!L2 1865.3166390713727
EVALUATED VALUE Sheet1!M2 3730.6332781427454
EVALUATED VALUE Sheet1!N2 7461.266556285491
EVALUATED VALUE Sheet1!O2 14922.533112570982
EVALUATED VALUE Sheet1!P2 29845.066225141964
EVALUATED VALUE Sheet1!Q2 59690.13245028393
EVALUATED VALUE Sheet1!R2 119380.26490056785
EVALUATED VALUE Sheet1!S2 238760.5298011357
EVALUATED VALUE Sheet1!T2 477521.0596022714
EVALUATED VALUE Sheet1!U2 955042.1192045428
EVALUATED VALUE Sheet1!V2 1910084.2384090857
EVALUATED VALUE Sheet1!W2 3820168.4768181713
EVALUATED VALUE Sheet1!X2 7640336.953636343
EVALUATED VALUE Sheet1!Y2 15280673.907272685
EVALUATED VALUE Sheet1!Z2 30561347.81454537
Evaluation done 0:00:00.052970
all done 0:00:00.622479
```

## xlcalculator

```
python xlcalculator_individual_SUM.py
loading file
model compiler made 0:00:00
INFO:root:File xl/sharedStrings.xml is not in archive.
read_and_parse_archive took 0:00:09.123948
build_code took 0:00:13.455735
now evaluating
evaluator made 0:00:13.456734
addresses made 0:00:13.456734
EVALUATED VALUE Sheet1!D2 12.502811628302023
EVALUATED VALUE Sheet1!E2 14.5727862427451
EVALUATED VALUE Sheet1!F2 29.1455724854902
EVALUATED VALUE Sheet1!G2 58.2911449709804
EVALUATED VALUE Sheet1!H2 116.5822899419608
EVALUATED VALUE Sheet1!I2 233.16457988392162
EVALUATED VALUE Sheet1!J2 466.32915976784324
EVALUATED VALUE Sheet1!K2 932.6583195356865
EVALUATED VALUE Sheet1!L2 1865.316639071373
EVALUATED VALUE Sheet1!M2 3730.633278142746
EVALUATED VALUE Sheet1!N2 7461.266556285492
EVALUATED VALUE Sheet1!O2 14922.533112570984
EVALUATED VALUE Sheet1!P2 29845.066225141967
EVALUATED VALUE Sheet1!Q2 59690.132450283934
EVALUATED VALUE Sheet1!R2 119380.26490056787
EVALUATED VALUE Sheet1!S2 238760.52980113574
EVALUATED VALUE Sheet1!T2 477521.0596022715
EVALUATED VALUE Sheet1!U2 955042.119204543
EVALUATED VALUE Sheet1!V2 1910084.238409086
EVALUATED VALUE Sheet1!W2 3820168.476818172
EVALUATED VALUE Sheet1!X2 7640336.953636344
EVALUATED VALUE Sheet1!Y2 15280673.907272687
EVALUATED VALUE Sheet1!Z2 30561347.814545374
Evaluation done 0:00:00.037979
all done 0:00:13.495712
```

## Koala2

```
python Koala2_individual_SUM.py
loading file
excel compiler made 0:00:02.794000
graph generated 0:24:24.278000
addresses made 0:24:24.279000
EVALUATED VALUE Sheet1!D2 12.5028116283
EVALUATED VALUE Sheet1!E2 14.5727862427
EVALUATED VALUE Sheet1!F2 29.1455724855
EVALUATED VALUE Sheet1!G2 58.291144971
EVALUATED VALUE Sheet1!H2 116.582289942
EVALUATED VALUE Sheet1!I2 233.164579884
EVALUATED VALUE Sheet1!J2 466.329159768
EVALUATED VALUE Sheet1!K2 932.658319536
EVALUATED VALUE Sheet1!L2 1865.31663907
EVALUATED VALUE Sheet1!M2 3730.63327814
EVALUATED VALUE Sheet1!N2 7461.26655629
EVALUATED VALUE Sheet1!O2 14922.5331126
EVALUATED VALUE Sheet1!P2 29845.0662251
EVALUATED VALUE Sheet1!Q2 59690.1324503
EVALUATED VALUE Sheet1!R2 119380.264901
EVALUATED VALUE Sheet1!S2 238760.529801
EVALUATED VALUE Sheet1!T2 477521.059602
EVALUATED VALUE Sheet1!U2 955042.119205
EVALUATED VALUE Sheet1!V2 1910084.23841
EVALUATED VALUE Sheet1!W2 3820168.47682
EVALUATED VALUE Sheet1!X2 7640336.95364
EVALUATED VALUE Sheet1!Y2 15280673.9073
EVALUATED VALUE Sheet1!Z2 30561347.8145
Evaluation done 0:00:00.015000
all done 0:24:24.295000
```

<!-- # Running it

From the root excel_evaluation_performance_tests directory
```python
python -m unittest discover -p "*_test.py"
``` -->

# TODO
- create a test which uses multiple worksheets.
