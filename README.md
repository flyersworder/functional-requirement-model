# Functional Requirement Model (FRM)

## About FRM

FRM is used in any hub or depot design (or redesign).
For the road transit hubs (RTHs), the inputs are truck, trailer, or any vehicle schedules, information including vehicle type, capacity, material type**, and arrival and departure times.

In the end, sort capacity, buffer size, and the number of doors are used for hub and depot design. Because the parameters are different per material type, the calculations are usually separate. Here is an overview of the whole process.

1. Inputs

    * Inbound schedule
    * Outbound schedule
    * Destination distribution
    * Operational parameters (e.g., uploading rate, offloading rate)

2. Steps

    * Schedule clean-up[^anote]
    * Volume availability (per 15 minutes or other time units)
    * Destination spread (time unit volume × destination distribution)
    * Sort simulation (linear programming, optimization)

3. Outputs

    * Sort window
    * Sort capacity
    * Buffer size
    * Door calculation


[^anote]:would be nice if we can have a first check to see whether the schedule makes sense and provide some warnings/alerts; also, schedules and destination distribution can come in different formats **the most distinguishing material type is conveyables (that can be processed with a sorter, i.e., autosort) and non-conveyables

## About this tool

### Background

Current FRM function is excel based, it is error prone and slow in caculation.
Qing Ye and Dewei Zhai work together to convert it into a self stand python based tool, in order to improve work efficiency.

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
