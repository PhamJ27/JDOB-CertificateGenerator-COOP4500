<h1> This is the John D' O Bryant Certificate/Awards Generator  </h1>
<br>

<!--    Explanation and Purpose     -->
<p>
    This project creates awards for each student based on a spreadsheet containing all the students and their designated certificate for the term. 
    The term is calculated based on the current time of generation and create a folder in the <strong>"output"</strong> folder.
</p>
<p> This project creates awards for each student using: </p>
<ul>
    <li> a template </li>
    <li> a spreadsheet - using: </li>
        <ul>
            <li>StudentFullName</li>
            <li>Year</li>
            <li>Certificate</li>
        </ul>
</ul>
<p>
    The program will create two outputs, a docx file and pdf file. 
    Use the docx file for last minute changes, but otherwise the pdf file works just as well for printing.
</p>
<br>

<!---    Requirements   --->
<p>
    This was programmed in Python 3.11.9 and in PyCharm Community Edition 2024.2.2.
    Although the requirements may be lower and regardless of IDE.
</p>
<p> This project requires the packages: </p>
    <ul>
        <li>python-docx</li>
        <li>docxcompose</li>
        <li>docx2pdf</li>
    </ul>
<p> It also uses some built-in modules such as: </p>
<ul>
    <li>csv</li>
    <li>datetime</li>
    <li>os</li>
    <li>pathlib</li>
</ul>

<!--    Required Files   -->
<p> The template MUST come in a form of a docx file, a sample is provided.  </p>
<p> The columns must be labelled as the above. The columns mean:            </p>
<ul>
    <li>StudentFullName - Student's Full Name in the order of: "FirstName LastName" </li>
    <li>Year - Year Of Graduation. Listed as a four digit number such as: 1999, 2005, 2067.</li>
    <li>Certificate - The specific award the student is granted. Example: High Honors, Honor Rolls, and  Honorable Mention.</li>
</ul>

<!---     How it Works    --->

<p>Example: </p>

|    StudentFullName | Year | Certificate       |
|-------------------:|------|-------------------|
|        Jotaro Kujo | 1987 | Honorable Mention |
|   Noriyaki Kakyoin | 1987 | Honor Roll        |
| Josuke Higashikata | 2001 | Honor Roll        |
|           John Rod | 2026 | High Honors       |

<br>

> The assign() - replaces the instance of placeholder text with the replacement text and appropriate formatting. 
> Examples of placeholder text being {{FULLNAME}}", "{{CERTIFICATE}}", "{{MONTH}}", "{{YEAR}}", etc.

> The readsheet() - opens the sheet and passes each row and year/month into the assign function.

> The merged versions of these functions are modified work in tandem with each other to create a singular docx file.

> filecheck() - prints the current working directory + checks if the spreadsheet and template are available and valid.

> yearCalc() - calculates the current grade level for the student based on Year of Graduation

> deTERMine() - determines the term based on given month.

> deTERMineSUFFIX() - determines the 1st, 2nd, 3rd, and 4th suffixes based on term.

> SchoolYearCalc() - calculates the school year range based on month.

> statcounter() - counts all students and counts for each award type: High Honors, Honor Roll, and Honorable Mention.

> all_paragraphs() - grabs the paragraphs in the document

> createCertificates - creates the certificate using a spreadsheet and given month and year.

<p> The main methods to create the pdf is createCertificates. </p>

<h2> Example Spreadsheet: </h2>
<img src=".\readMEAsset\SheetExample.png">

<h2> Example Template: </h2>
<img src=".\readMEAsset\HPTemplatePreview.png">

<h2> Example Final Product: </h2>
<img src=".\readMEAsset\MakotoNaegiExample.png">