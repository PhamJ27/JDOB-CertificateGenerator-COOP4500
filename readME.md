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

    <!--    Requirements   -->
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
<br>

    <!--     How it Works    -->
<p>
    The program also has some statistical functions such as: statcounter() and fileCheck().
    statcounter() counts all students and counts for each award type:
</p>
<ul>
    <li> High Honors </li>
    <li> Honor Rolls </li>
    <li> Honorable Mention </li>
</ul>
<p>
    The assign() function replaces the instance of placeholder text with the replacement text and appropriate formatting. 
    Examples of placeholder text being {{FULLNAME}}", "{{CERTIFICATE}}", "{{MONTH}}", "{{YEAR}}", etc.

    The readsheet function opens the sheet and passes each row and year/month into the assign function.

    The merged versions of these functions are modified work in tandem with each other to create a singular docx file.

    
</p>
<p></p>