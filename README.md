# diff-doc-osx

A partial port of TortoiseSVN/TortoiseGit diff-doc.js to Objective-C on Mac OSX using ScriptingBridge.
For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/diff-doc.js.

The code is distributed under the GNU General Public License. 

## Prerequisites

Microsoft Word needs to be installed. Contrary to the source script, this port
does not support using OpenOffice (or LibreOffice) to perform the comparison. 

## Usage 

`diff-doc-osx <absolute-path-to-base.doc> <absolute-path-to-new.doc>`


## Build instructions

Open and build the project in XCode. The header file `Word.h` was initially generated
using the follwing command (see also this documentation from [Apple](https://developer.apple.com/library/content/documentation/Cocoa/Conceptual/ScriptingBridgeConcepts/UsingScriptingBridge/UsingScriptingBridge.html))

`sdef /Applications/Microsoft\ Word.app | sdp -fh --basename Word`

The resulting file was adapted so as to be free of compile errors (but not free of warnings).

## Known issues

* Relative paths to documents are not supported.
* Newer versions of Microsoft Office apps are sandboxed. This leads to the annoying
"Grant File Access" dialog to pop up for each of the documents to be compared in cases
where Word does not have permission to access the respective file already.

## Future work

Create a formula for [`brew`](https://github.com/Homebrew).
