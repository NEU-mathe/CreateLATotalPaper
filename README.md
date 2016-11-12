# CreateLATotalPaper

A Tool for Generating Total Question Bank of Linear Algebra Unit Quiz System and College Mathmatics Learning-in-process System as Doc.

This program is part of the tool kit NEU Mathe.

## Usage

This program is designed for developers at present. You need to rebuild after changing some arguments.

Some useful arguments are at line 76~79, Program.cs:

+ const int chapter = 6;
+ const int zNoCount = 5;
+ const int tNoCountL = 100;
+ const int tNoCountR = 172;

The variable _chapter_ shows the chapter you would like to generate as Doc.
The variable _zNoCount_ with a value 5 shows that each question has one stem and four branches.
The variable _tNoCountL_ and _tNoCountR_ shows that the number range of the questions saved in the local directories. You need to download the question bank from ftp, whose username and password can be got by either packet capture or reflection.

## Contribute

This series of tools do not have a specified person or team to maintain, developers ususally spare no time on it since they don't use it anymore. Thus, we are in an urgent need of your contribution. Contribute by fork and pull request to promote human emancipation. Thanks and have a good day.