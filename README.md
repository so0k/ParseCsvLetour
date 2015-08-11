## Introduction

This is a proof of concept PowerShell 3.0 script to parse HTML contained within a sample CSV file.

Powered by:

* [nuget](http://www.nuget.org)
* [HtmlAgilityPack](https://www.nuget.org/packages/HtmlAgilityPack)

Original gist: https://gist.github.com/so0k/06425d791c9afcc40ead

## Quickstart

To run from powershell command line:

```powershell
irm https://raw.githubusercontent.com/so0k/ParseCsvLetour/master/Parse.ps1 | iex
```

* irm is an alias for `Invoke-RestMethod`
* iex is an alias for `Invoke-Expression`

**Note**: Never pipe random scripts from the internet to your machine without knowing what they do.. you can review the script before running the line by going to the url: https://raw.githubusercontent.com/so0k/ParseCsvLetour/master/Parse.ps1 in your browser first.

**Note**: If you want to download the script and execute the script separately, you'd need to ensure your execution policy is set as follows (this will require Administrator elevated rights):

```powershell
Set-ExecutionPolicy RemoteSigned 
```

## How does the script work?

The powershell script will do the following:

* Prompt the user for an input csv file using an `OpenFileDialog`.
* Test to see if Output.csv already exists in work direcotry. If Output.csv already exists, the script will pause. User may abort the script with <kbd>CTRL</kbd>+<kbd>C</kbd> or press <kbd>ENTER</kbd> to continue. If the user chooses to continue, Output.csv will be overwritten.
* Test to see if nuget is installed on the system. If nuget is not installed, the script will look for it within the work directory. If nuget is not found within the work directory, nuget will be downloaded via https://www.nuget.org/nuget.exe.
* Use nuget to download HtmlAgilityPack version 1.4.9 locally
* Read the csv file to Memory (ensure sufficient memory is available)
* Load the HTML field into an `HtmlDocument`
* Use XPath to select all `div` elements with a `classname` of `desc`
* Create new csv records only holding the HTML fragment of the previously selected div elements
* Write the collection of csv records to disk

## Demo

![screencast.gif](screencast.gif "screencast")

[reddit thread](https://www.reddit.com/r/excel/comments/3gfae8/extracting_specific_html_code_from_within_a_cell/)

## License

This code is released under the GNU General Public License (GPL).

The GPL is a copyleft license that requires anyone who distributes this code or a derivative work to make the source available under the same terms.
