Excel formatting automation
===========================

shell/command line application which produces a formatted Excel file

    usage: xlsxfmt source format [options] [output]
    
    source (Excel file)
    format (yaml formatting description file)
    [output] output file name (overrides file name determined from values in format and options)
    
    Options:
    --output-filename-prefix=<prefix>		prefix to be added to the beginning of the output file name

    --output-filename-postfix=<postfix>		postfix to be added to the end of the output file name (before the extension)

    --grand-total-prefix=<prefix>			string to be prepended to the "Grand..." on the last line of any subtotaling

	--burst-on-column=<sourceColumnName>	If specified, each sheet in source must contain a column named <sourceColumnName>
											Each unique value (across all sheets) in this column will result in a separate
											output file. The name of each resulting output file will have that value included
											in its name immediatly after any output-filename-prefix.    

## formatting YAML specification 

### format
General/global settings for this format.

* **name**
* **description**
* **version**
* **logo-path**
* **output-filename-base**
	* This will be surrounded with any pre/postfixes specified as command line options.
* **burst-on-column** -- sourceColumnName
	* If specified, each sheet in source must contain a column named sourceColumnName.
	* Each unique value (across all sheets) in this column will result in a separate output file. 
	* The name of each resulting output file will have that value included in its name immediately after any output-filename-prefix.
* **empty-sheet**
	* **exclude** -- if sheets with no data should be excluded from output (default: false)
	* **default-text**
		* value to be entered in the cell where the first value would be (usually A2)

### defaults
Default settings. The available scopes for defaults are **sheet**, **column**, and **font**.

### sheet
List of one or more sheets to be created in the output file. Sheets will appear in the output file in the order listed.

* **name**
	* The name of the sheet created in the output file.
* **source**
	* The name of the sheet from the source file.
	* The same source sheet may be used to generate more than one sheet in the output.
	* Defaults to the same as **name**
* **freeze-on-cell**
	* Cell on which to freeze panes.
* **header-row-bgcolor**
	* Background color to be applied to the header (A) row.
	* Specified as a hex value.
* **grand-total-row-bgcolor**
	* Background color to be applied to the "Grand..." row of subtotaling.
	* Specified as a hex value.
* **sort** -- list of columns to sort on (in the order the sort will be applied)
	* **column** -- the name of the column
	* **direction** -- the direction of the sort
		* permitted values: ascending, decending
		* default: ascending
* **top-n-rows** 
	* number of rows to include in output
	* removal of excluded rows shall be done after all sorting and other removals (such as those based on column.stop-values) have been completed 
* **hidden** -- if sheet should be hidden (default: false)
* **include-logo** (true/false)
	* If true, logo will be inserted at the top left corner of the sheet and the height of the header row adjusted to show the header row text just below the logo. 
* **totals-calculation-mode** (formula/internal)
	* formula - subtotals will be calculated in Excel (by inserting a SUBTOTAL function)
	* internal - subtotals will be calculated in xlsxfmt and the result transferred to the output cell
* **column**
	* List of one or more columns to be created/formatted in the output sheet. See below for the possible sub-keys.



### column

* **name**
	* The output column name.
* **source**
	* The name of column in the source file/sheet.
	* The same source column may be used to generate more than one column in the output.
	* Defaults to the same as **name**
* **stop-values**
	* List of one of more values which, if equal to the value of a cell in this column, will result in that row being excluded from the output. These exclusions shall be processed before any formatting, grouping, totaling, or formulas are added. 
* **required-values**
	* Only rows from the input which contain one of the values in the supplied list for this column shall be included in the output. This shall be processed before any formatting, grouping, totaling, or formulas are added.
	* This list of values may include both strings and numbers. 
* **width**
	* permitted values: "auto" or numeric
	* default: auto
	* If "auto" column will be autofix after all other formatting is applied.
* **format-type**
	* permitted values: GENERAL, ACCOUNTING, NUMBER, TEXT, DATE
	* default: GENERAL
* **decimal-places**
	* Applies to NUMBER and ACCOUNTING format-types.
* **date-format**
	* Excel valid formatting string
* **hidden** -- if column should be hidden (default: false)
* **totals-calculation-mode** (formula/internal)
	* formula - subtotals will be calculated in Excel (by inserting a SUBTOTAL function)
	* internal - subtotals will be calculated in xlsxfmt and the result transferred to the output cell
* **conditional-formatting**
	* **type**
		* permitted values: databar
	* **style**
		* permitted values: gradient-blue, gradient-green, gradient-red, gradient-orange, gradient-ltblue, gradient-purple
	* *additional options to be added in the future* 
* **subtotal**
	* **group**
		* insert a subtotal line for each change in value of field
		* when multiple fields have this option set, they are to be applied left to right such that expanding the highest numbered grouping level expands the rightmost subtotals
		* this defines the grouping, not the field to actually subtotal which is defined by the column with the function setting
	* **total-row-bgcolor**
		* (hex) color to set the background to on the rows containing subtotals based on the change in this column
	* **function**
		* subtotal function to apply to this field
		* value is one of the [function\_num](https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939) values allowed for the SUBTOTAL function
		* should not be used on the same field as subtotalOnChange

### font
May be used as a child of either **sheet** or **column**. When used for both, the more specific wins (i.e. a column level setting overrides a sheet level setting). Any value not specified will use the Excel/system default.

* **family**
* **size**
* **style**
	* permitted values: bold, italic, underline, none
	* May contain a list of multiple values.
* **header**
	* Font settings to be applied specifically to header rows.
* **data**
	* Font settings to be applied specifically to data (non-header/footer) rows.
* **footer**
 	* Font settings to be applied specifically to footer rows (i.e. subtotal rows).

