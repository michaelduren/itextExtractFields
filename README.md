Extracts fields from a PDF file using iTextSharp

Copyright (C) 2018, Cyber Moxie, LLC

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

 
This script extracts fields or values from a PDF file where values have the format:
         
         <Label>: <Value>

and prints the data to the screen in CSV format.  The process allows you to map the values in the file to a column 
in the CSV file by specifying a column name in the field definition field.  The enable you to extract values and combine
them into a single field.  This is helpful for address values were address, address2, city, state, zip are separate fields in the file, but
should be combined into a single column field in the output.  First name, last name is another example.  THIS ONLY WORKS FOR consectutive 
values in the file.  A future enhancement could sort the tuple below to combine fields that are dispersed.

 
Parameters:
    fieldLabelsFile - file containing field labels of values to search for and extract
    inputFile - file that is target of extraction

The fieldsLabelsFile has the following format:
  
  fileLabel1:,ouputLabel1
  fileLabel2:,ouputLabel2

If the outputLabel of consecutive fields match, then the output will combine these fields into a single column.  See the examples in this 
project.
