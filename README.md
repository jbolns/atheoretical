# PapayaWatermelon, VBA projects of various sorts.
This repo currently contains two projects. Both projects are in very early stages. Expect bugs. Loads of them.

## 1. Atheoretical 
**Auto (supervised) ML of sorts**. A set of VBA scripts to rapidly produce supervised ML analyses in Excel. Helpful when you run need to run regressions/supervised ML in Excel on a periodical basis and you want to save yourself some time. [Link to folder](https://github.com/jbolns/papayawatermelon/tree/main/Atheoretical).

### Notes
- The scripts are given as *.bas* files, which can be imported easily into Excel's VBA console. 
- The file is currently limited to clean datasets with quantitative continuous variables.
- Alternatively, the repo contains an Excel document with preloaded scripts and pre-populated (dummy) hyper-parameters, for testing.
- The dataset in this folder is also for testing purposes. 

### Pending:
- [ ] Documentation. The Excel file includes guidance, but more detailed documentation would not hurt.
- [ ] Stress testing. The file currently works fully with the sample dataset, but bugs are incredibly likely.
- [ ] Code cleaning. There are still macros that use repeated code sequences.

## 2. WRD to HTML
**Document type conversion**. A VBA script to convert (simple) Word documents into HTML. Helpful when you write a lot in Word (e.g., op-eds, journal articles, data analysis briefs) and need to frequently upload to HTML-driven websites. [Link to folder](https://github.com/jbolns/papayawatermelon/tree/main/WRD%20to%20HTML).

### Notes
- The script is given as a *.bas* file, which can be imported easily into Word's VBA console. 
- Alternatively, the repo contains a Word document with the script preloaded and dummy text, for testing.
- The picture in this folder is also for testing.

### Pending:
- [ ] Documentation. Currently, you would need to read the VBA code to understand the script's capacities and limitations, and how it works.
- [ ] Images. The script currently supports image conversion, but this aspect of the script could be improved significantly. 
- [ ] Tables. Currently, the file does not support table conversion. 
- [ ] Markdown support. Once all kinks are resolved, it should be possible to add an option to convert to Markdown.

## Usage/Licensing
I honestly have not thought about this just yet. I *really* do not think the files are yet at production or even pilot-testing level. So, for now, let's say it's all rights reserved.
