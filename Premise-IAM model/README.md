# E-waste-Challenges-of-Generative-Artificial-Intelligence
This is the program to realize global server E-waste prediction algorithm presented in paper 'E-waste Challenges of Generative Artificial Intelligence' The functionality of this program is to calucalte (predict) the number of servers with given training/inference scale parameters in the configurated period. The result contains mean value, standard deviation and cumulative standard deviation from the starting time frame.

1.System requirements
The program is written and tested in Python3.9, win64 version. Any Python 3 version might be compatible with this program.
The following Python libraries should be installed properly as well:
 - numpy
 - xlrd
 - pandas
 - openpyxl

2.Installation guide
Please install Python environment (preferably 3.9 or later version) and the Python libraries listed in System Requirements. The installation process follows standard Python installation steps and takes only a few minutes to complete.

3.Demo
The program comprises a 'Server prediction.py' Python script and a 'Demo data.xls' excel file. Configure the excel file to settle the scenarios and excute Python script with Python interpretor, the result file will be generated in the current folder. Detailed steps:

-A. Open 'Demo data.xls', where you can find 2 sheets. Sheet 'Input configuration' allows to change the value of algorithm required parameters: 'Number of random iterations, Average lifespan of server, Training time for one GAI model, Number of GPU per server, Upgrade strategy, Tokens required per user per day'. Please see detailed explaination of these parameters in Supplementary Information. Sheet 'Computation data' contains future (predicted) data to support server number calculation, including 'Number of parameters per GAI model, Number of training data/token, Compuational power for training, Compuational power for inference, Global GAI user number, Efficiency for training, Efficiency for inference, Sparsity rate' in the first 8 columns. The rest columns list out data in other scenarios for the reference. You can copy those data into first 8 columns to change scenario setting. The default value in the sheet are set as a demo for baseline scenario. You can change any parameters to change the scenario setting. In the third sheet, a probability table following normal distribution is presented, to calculate the retirement of servers within a time range.

-B. Once 'Demo data.xls' is well set, run 'Server prediction.py' with installed Python interpretor. Wait a few seconds and the result file (named 'Result.xlsx') will be generated in the same folder.

-C. Open 'Result.xlsx'. It contains 7 columns. Column A refers to time index (by default, at seasonal interval). Column B refers to the mean value of input server number estimation at this time frame. Column C refers to the standard deviation of column B at this time frame and column D refers to cumulative strandard deviation for column B from index 0 to current time index. Column E refers to the mean value of output (retirement) server number estimation at this time frame. Column F refers to the standard deviation of column E at this time frame and column G refers to cumulative strandard deviation for column E from index 0 to current time index. 

4.Instruction for use
You can change the values of parameters in 'Demo data.xls' to compute the time-sequenced server generation amount in configured scenario, by following the three steps illustrated in previous section.
In addition, the result of scenarios presented in the manuscript can be reproduced using the values listed in the excel sheet (in sheet 'Compute data', columns after 'Alternative data') or in the Supplementary Information (See Table S1, Figure S2, Figure S3 and Figure S4). You can substitute the configuration values with these suggested values then run the Python script.
