# VBA-VBS

The most of the funcions here works both as VBA and to run as VBS it may need to realize some changes.

## ConvertChar:
This function receives an integer n as input and returns a string with the corresponding character in the range from "A" (0) to "ZZ" (701).

## DrawCircle_MSWord:
The DrawCircle function is used to draw a circle in a Microsoft Word document. The function includes and formats text within the circle, as well as formatting the shape of the circle itself.

The function takes no input arguments, however it can be changed to receive as input the text, color, or other desired properties. When called, it will create a circle with the specified text, font color, line weight, dash style, style, transparency, and fore and back colors in the current Microsoft Word document.


## LevenshteinDistance:
This function calculates the Levenshtein distance between two strings. The Levenshtein distance is a measure of the similarity between two strings, defined as the minimum number of single-character edits (insertions, deletions or substitutions) required to change one string into the other.

### Usage:
`Levenshtein("string1", "string2")`

The function will return an integer representing the Levenshtein distance between the two strings.

### Reference:
[StackOverflow](https://stackoverflow.com/questions/4243036/levenshtein-distance-in-vba)

## ReplaceSpecialCharacters
The ReplaceSpecialCharacters function removes special characters from a string. It takes in a single string as an input and returns a string with all the special characters removed. The function uses a loop to iterate through each character in the input string, and if the character is not a letter or a number, it is replaced with a blank space.

### Reference:
[Microsoft Community](https://answers.microsoft.com/en-us/msoffice/forum/all/removing-special-characters-with-regular/d62d50b7-8586-4f08-ac7d-c5212929074a)