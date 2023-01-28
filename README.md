# matlab2word
A MatLAB library that allows to write your calculations into Word template file.
It might be useful when you have a report template which has specific formatting and requires some calculations.
See how to use this library in files matlab2word_example_en.mlx or matlab2word_example_ru.mlx (open these files in MatLAB).


Version history

v1.0.0 - Initial version.
-   Abilities to write text, numbers and plots into Word

v1.1.0 - Added some features
-   Now your result saves automatically by adding _out part to your input file.
If you want to select a place to save, use method .SaveManually()
-   You can change your default imaginary unit using function .SetImaginaryUnit('j')
-   You can use data for replacement not only from MATLAB, but also from other file
