# PowerBuilder-Dw2Doc-Example

## Dw2Excel

This application converts supported DataWindow presentation styles into Excel Spreadsheets using [NPOI](https://github.com/nissl-lab/npoi) to write the XLSX file.

### Features

Allows converting Grid, Freeform, Tabular DataWindows to Excel files while attempting to preserve the visual appearance as faithfully as possible.

Supports converting most commonly used DataWindow objects and their properties (also supports expressions). 

### Usage

#### Convert DataWindow to Excel

Run the `dw2excelapp` target. Select a DW style you wish to export. A preview of the DataWindow will be displayed. You can filter the data by enabling the `Filter data`, change the export settings by clicking on the gear (⚙️) button or begin the export process by clicking the `Save as XSLX` button.

#### Filtering 

You can filter the data presented in the DataWindow by clicking the `Filter data` checkbox and selecting a Column, Operator and entering a value. This will only show the data that satisfies the configured condition.

#### Options

On the main window if you click on the gear ⚙️ button you will access the *Conversion settings* window, in which you're able to configure the following parameters:

- Thresholds: Objects offset by an amount less than X/Y threshold will be considered to be on the same row/column
- Misaligned objects: Objects whose position is greater than the threshold, but overlap are converted into floating objects (e.g. columns get converted into Excel text boxes). This option omits exporting these objects.
- Bands: Select the bands you want to include in the conversion.
- Conversion Targets: Select the object types you want to include in the conversion.
- Property targets: select the properties you want to include in the conversion. 

#### Save as XLSX

##### New File

If the selected location does not exist, a new file will be automatically created for you.

##### Existing File

If the selected location points to an already existing file you will be prompted whether you want to overwrite the file, or append to it.

- Overwrite: Replaces the existing file. 
- Append as a sheet: You will be asked to input the name of the sheet that will be created.

#### Deploy along with your PB project

All the code for the conversion is contained inside the *dw2doc.pbl* library with dependencies on *csharpextensions.pbl* and *common.pbl*. Make sure to include all three of them if you plan on using it in your projects. The main entry point for the conversion is *n_cst_dw2excelconverter.of_convert* function.

### Algorithm

- Determine the amount and height of the DW's bands.
- Iterate over the controls and determine the band that contains them.
- Iterate over each band and control and define a row layout that fits the controls and bands. There are some DWOs that cannot be put inside cells (bitmap, line, geometric shapes), these are created as shapes and anchored to the closest cell. If it's impossible to put two objects into their own row, one of them is converted to a Text Box shape and anchored to the closest cell.
- Iterate over each control and define a column layout that fits the controls and bands. If two controls overlap on the X axis, it will attempt to create additional columns and merge them to accommodate them. After this step the application created a `VirtualGrid` that defines the objects positions' in terms or rows and columns.
- The application then iterates over each DW row, and using the created `VirtualGrid` maps each row value's to a control building an XSSFWorkbook.
- Then it saves this XSSFWorkbook to disk.

### Limitations

The algorithm is on its early stages and currently has a very narrow support scope. 

#### Themes

Themes are not supported, due to the fact that their effects are applied after the DataWindow properties have been evaluated, thus their colors cannot be read through `DataWindow.Describe`. 

#### Supported DataWindow styles

- Grid
- Tabular
- Freeform

#### Supported DW Objects

- Column
- Compute
- Text
- Bitmap
- Lines
- Rectangle
- Oval
- Button

#### Object property support

The demo currently only translates the object properties listed in the following table. Please note that some properties don't apply to all controls.

| Object/Property  | Column | Compute | Text | Bitmap | Line | Rectangle | Oval | Button |
| ---------------- | ------ | ------- | ---- | ------ | ---- | --------- | ---- | ------ |
| Visible          | ✅      | ✅       | ✅    | ✅      | ✅    | ✅         | ✅    | ✅      |
| Alignment        | ✅      | ✅       | ✅    |        |      |           |      |        |
| Color            | ✅      | ✅       | ✅    |        |      |           |      |        |
| Background.Color | ✅      | ✅       | ✅    |        |      |           |      |        |
| Font.Face        | ✅      | ✅       | ✅    |        |      |           |      |        |
| Font.Height      | ✅      | ✅       | ✅    |        |      |           |      |        |
| Font.Weight      | ✅      | ✅       | ✅    |        |      |           |      |        |
| Italics          | ✅      | ✅       | ✅    |        |      |           |      |        |
| Underline        | ✅      | ✅       | ✅    |        |      |           |      |        |
| Brush.Color      | ✅      | ✅       | ✅    |        |      | ✅         | ✅    |        |
| Brush.Hatch      |        |         |      |        |      | ✅         | ✅    |        |
| Pen.Color        |        |         |      |        | ✅    | ✅         | ✅    |        |
| Pen.Style        |        |         |      |        | ✅    | ✅         | ✅    |        |
| Filename         |        |         |      | ✅      |      |           |      |        |
| Transparency     |        |         |      | ✅      |      |           |      |        |
| Text             | ✅      | ✅       | ✅    |        |      |           |      | ✅      |



This demo has been tested with a very small set of DataWindows so when presented with other, it might produce unexpected results; this is currently a work in progress.



## Dw2Word

Exports a DataWindow into a Word Document using templates.

### Features

Exports a DataWindow record into a Word document by opening a template and replacing placeholder text with the value of the columns. This approach only works for writing one record at a time.

### Usage

Run the `dw2wordapp` target, click the `Use template` check box and select the template you want to use. Then click the `Save as Word` button.

#### Modifying the templates

The templates contain placeholders with the column's names. If you want to adapt the templates to fit your own DataWindows, make sure that the column names are enclosed in brackets `{}`.
If you want to create your own templates, create a new document and format it however you wish. Add placeholders for the columns.

### Limitations

Currently DataWindows can only be exported to Word documents via templates that need to be tailored for each DataWindow (it's necessary to position the columns manually in the document).

## Feedback

If you run into any unexpected errors that prevent this demo from running, or would like to submit a suggestion for a feature you would like to see added to the demo, submit a ticket on our [Support Portal](https://www.appeon.com/standardsupport/) so that we can follow up on it.
