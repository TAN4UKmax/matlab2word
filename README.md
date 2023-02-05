# matlab2word
A MATLAB library that allows writing your calculations into a Word template file.
It might be useful when you have a report template that has specific formatting and requires some calculations.
See how to use this library in files matlab2word_example_en.mlx or matlab2word_example_ru.mlx (open these files in MATLAB).
If you still have questions, feel free to contact me: tan4ukmak7@gmail.com
Special thanks to Оль Роман and Шевченко Алина for help in developing this library.


Version History

v1.0.0 - Initial version.
-   Abilities to write text, numbers, and plots into Word.

v1.1.0 - Added some features
-   Now your result saves automatically by adding _out part to your input file.
If you want to select a place to save, use the method .SaveManually().
-   You can change your default imaginary unit using a function .SetImaginaryUnit('j').
-   You can use data for replacement not only from MATLAB but also from another file.

v1.2.0 - Code improvements and added table replacement
-   Now for using comma as a decimal separator call .SetDecimalSeparator(',').
-   You can paste data in a table with an undefined number of rows.

v1.3.0 - Spelling fix and saving file features
-   You can specify output file name explicitly as well as for the input file.
