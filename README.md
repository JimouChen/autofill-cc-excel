# autofill-cc-excel

An app based on the Windows platform that helps CC (myGF) to automatically fill out excel forms

<p align="center">
  <a href="https://github.com/JimouChen/autofill-cc-excel">
     <img src="https://github.com/JimouChen/autofill-cc-excel/tree/main/images/st.png" alt="Logo" width="80" height="80">
  </a>

<h3 align="center">Autofill</h3>
  <p align="center">
    Manpower-saving App, constantly improve and update App according to demand
    <br />
    <a href=""><strong>Documentation</strong></a>
    <br />
    <br />
    <a href="https://github.com/JimouChen/autofill-cc-excel/blob/main/README.md">Readme-en</a>
    ·
    <a href="https://github.com/JimouChen/autofill-cc-excel/blob/main/READMECN.md">Readme-cn</a>
  </p>
</p>

<!-- PROJECT SHIELDS -->

![platform](https://img.shields.io/badge/platform-win64-lightgrey.svg)
![](https://img.shields.io/badge/License-Apache%202.0-green.svg)
![](https://img.shields.io/badge/autofill--cc--excel-v1.0-blue.svg)

<br />

## Notices

- This software can only be used privately and cannot be used for commercial use
- Since the data involves company-related information, the data format is inconvenient to display and can only be used
  privately


## Setup
###### Pre-development configuration requirements

- Python >= 3.8

###### **Installation Steps**

- Clone the repo

```sh
git clone https://github.com/JimouChen/autofill-cc-excel.git
```
- Install packages

```shell
pip install -r requirement.txt
```
- Run or Package into exe 
```shell
python help_cc_excel.py
```
or
```sh
Pyinstaller -F -i st.ico -n exec_file_name help_cc_excel.py
```
> and then double-click to run the generated exe file in the dist directory

### File Directory Description

```
.
├── LICENSE
├── README.md
├── help_cc_excel.py
├── images
│   └── st.png
├── requirement.txt
└── st.ico

```


### Version Control

- The project uses Git for version management

### Author

author@JimouChen/Neaya
