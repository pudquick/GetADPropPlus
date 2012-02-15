#GetADPropPlus - A modification of an Excel macro for looking up AD properties

GetADPropPlus is an extension of an Excel macro I found here: http://www.remkoweijnen.nl/blog/2007/11/01/query-active-directory-from-excel/

It can let you look up AD attributes from within Excel, which can help immensely with spreadsheets relating to computer or user objects. Usage of the function is as follows:

`=GetADPropPlus(attributeToMatch, objectType, matchString, attributeToReturn)`

An example usage for returning the display name of a user, based on their login account, would be:

`=GetADPropPlus("samaccountname", "User", "<login name here>", "displayname")`

This will perform an AD search of the login name provided and attempt to find a match for a "User" object with that login name value for their samaccountname attribute. If found, it will return the displayname attribute for that object.

An example usage for returning the description of a computer object, based on the computer's domain name, would be:

`=GetADPropPlus("cn", "Computer", "<computer name here>", "description")`

This will perform an AD search of the computer name provided and attempt to find a match for a "Computer" object with that value for their cn attribute. If found, it will return the description attribute for that object.

As this is an Excel function, when used you can replace the "<whatever name here>" strings with cell addresses instead and it will perform the lookup using the value of the cell in the Excel spreadsheet.

##Credits

- GetADPropPlus is written by pudquick@github 

##License

GetADPropPlus is released under a standard MIT license.

	Permission is hereby granted, free of charge, to any person
	obtaining a copy of this software and associated documentation files
	(the "Software"), to deal in the Software without restriction,
	including without limitation the rights to use, copy, modify, merge,
	publish, distribute, sublicense, and/or sell copies of the Software,
	and to permit persons to whom the Software is furnished to do so,
	subject to the following conditions:

	The above copyright notice and this permission notice shall be
	included in all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
	EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
	MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
	BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
	ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
	CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.


