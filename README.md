# Functional Requirement Model (FRM)

## About FRM

FRM can be used in any hub or depot design (or redesign).
For the road transit hubs (RTHs), the inputs are truck, trailer, or any vehicle schedules, information including vehicle type, capacity, material type (most noticeably conveyables that can be processed by a sorter and non-conveyables that are usually handled by forklifts), and arrival and departure times.

In the end, sort capacity, buffer size, and the number of doors are used for hub and depot design. Because the parameters are different per material type, the calculations are usually separate. Here is an overview of the whole process.

1. Inputs

    * Inbound schedule
    * Outbound schedule
    * Destination distribution
    * Operational parameters (e.g., uploading rate, offloading rate)

2. Steps

    * Schedule clean-up<sup>1</sup>
    * Volume availability (per 15 minutes or other time units)
    * Destination spread (time unit volume × destination distribution)
    * Sort simulation (linear programming, optimization)

3. Outputs

    * Sort window
    * Sort capacity
    * Buffer size
    * Door calculation


<sup>1</sup>would be nice if we can have a first check to see whether the schedule complies to the business rules and provide some warnings/alerts; also, schedules and destination distribution can come in different formats 

## About this tool

### Background

Current FRM function is excel based, and it is error prone and slow in caculation.
We aim to convert it into a self stand python based tool in order to improve work efficiency and robustness.
Ultimately, we want to design a GUI and package the module as an executable file so that it can be easily distributed.

### Structure of code

**src** folder place source code.
**tests/unittests** place testcases based on pytest.
**main.py** utilize the tool.

The tool works as a python module inside a main function.

Use this to check code quality before pushing

```bash
pylint src
pytest tests/unittests
pip install .
```

### Way of working

Don't update on master branch directly.
Create a branch based on Jira task number.
Use push request to merge into master.
