# packetizer
Convert question docs + master spreadsheet to category-randomized packets

<h3>Requirements:</h3>
A completed quizbowl tournament, consisting of a set of category Google Docs containing all the questions in a category.
A spreadsheet containing two tabs that have the packetized tossup and bonus answerlines. See "Answerline formatting guidelines" for details.
Spreadsheet tabs should look something like this (spoilers for the first two rounds of CMST II):
<img width="550" alt="sheet_formatting" src="https://user-images.githubusercontent.com/8041675/230182992-2212174c-db09-49be-bf90-57f20cd35f37.png">

In particular, the first column MUST contain all of the categories, formatted as "Category - Subcategory". The category MUST match the corresponding string in the dictionary of document URLs in the script (lines 2-7, see below).

<h3>Usage:</h3>
1. Open tournament master spreadsheet.

2. Go to Extensions -> Apps Script.

3. Add a file, name it whatever you want, paste the code in generatepackets.gs.

4. Change the following things in the script:

-links to the question docs (lines 2-7)

-set extraRow to true if the sheet has an extra row before the start of the questions (e.g. the CMST II sheet). (line 9)

-set writerTags to false if the questions do not have an extra line under the answerline saying who wrote it, what category it's in, etc. (line 10)

-names of the sheets containing the packetized TU and bonus answerlines (lines 16-17).

-the information printed at the top of each packet (lines 36-40).

5. Run the script. This should generate packets. Missing answerlines will be highlighted in the spreadsheet and noted in the packet. If at least one answerline is missing from the packet, it will have "INCOMPLETE" appended to its title.

<h3>Answerline formatting guidelines:</h3>

answerlines in the sheet must match either the beginning of an answerline in the question doc or the exact primary answer, that being all of the text in the question doc answerline that is bold and underlined before the first square bracket. The script will attempt to fix capitalization mismatches.

For example, if the answerline in the document is "ANSWER: quantum <b>Hall</b> effect [accept <b>Hall</b> effect]", the script will work if the answerline in the spreadsheet is "Hall" or "quantum Hall" but NOT "Hall effect".

<h3>Limitations:</h3>

-Currently generates packets in the user's root folder, rather than a tournament folder.

-No way to set the RNG used in generating the order of questions in packets, so results will be different on each run.

-Currently, a order of questions is valid if questions in the same category are not consecutive, the category of a tossup and its corresponding bonus are different, and two science/history/literature appear in each half. If you have a better way of generating valid tossup/bonus orderings, feel free to replace the corresponding functions and/or let me know and I'll update the script.
