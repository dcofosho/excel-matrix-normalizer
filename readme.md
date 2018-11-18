<h1>Python Excel Matrix Normalizer</h1>


<p>
This code takes an Excel spreadsheet that looks like this:

!["Example"](https://github.com/dcofosho/excel-matrix-normalizer/blob/master/example.png)


...And turns it into a spreadsheet which looks like this:

!["Example"](https://github.com/dcofosho/excel-matrix-normalizer/blob/master/output.png)

<ul>
<br>
Prerequisites
	<li>You will need a command line terminal from which to run the program. I recommend Git Bash (https://gitforwindows.org/)</li>
	<li>You will need Excel (https://products.office.com/en-us/excel?legRedir=true&CorrelationId=bf0cede6-79fc-46d2-bd7d-1bc52f2703ca&rtc=1).</li>
	<li>Download Python 3(https://www.python.org/downloads/). If you're not sure if you have it already, you can check by running <code>python -v</code> from your command line terminal.</li>
	<li>This code requires the <code>openpyxl</code> Python library. Install it with the command <code>pip install openpyxl</code></li>
</ul>

<ul>
How to run
	<li>Download or clone the repository to your local machine.</li>
	<li>In your command line terminal, navigate to the project folder using cd.</li>
	<li>Enter the command <code>python normalize.py [source file name] [source sheet name] [target file name] [start column] [end column] [start row] [end row]</code>
	<li>For example, to run the example provided, use the following command: <br> <code> python normalize.py example.xlsx Sheet1 output.xlsx A K 2 12
	</code>
	</li>
</ul>

