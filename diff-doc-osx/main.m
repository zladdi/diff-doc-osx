//
//  main.m
//  diff-doc-osx
//
//  A partial port of TortoiseSVN/TortoiseGit diff-doc.js to Objective-C on Mac OSX using ScriptingBridge.
//  For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/diff-doc.js
//
//  The OpenOffice portion was not ported.
//
//  Word.h header had to be adapted after having been generated. This is a known issue, see e.g.
//    http://stackoverflow.com/questions/15338454/scripting-bridge-and-generate-microsoft-word-header-file
//
//  The Objective-C code was partly inspired by this forum post:
//    https://discussions.apple.com/thread/2623068
//
//  This file is distributed under the GNU General Public License.
//
//  Author: Zlatko Franjcic

#import <Foundation/Foundation.h>
#import "Word.h"

int main(int argc, const char * argv[]) {
    @autoreleasepool
    {
        if(NSApplicationLoad())
        {
            WordApplication * word;
            NSString *sTempDoc, *sBaseDoc, *sNewDoc;
            WordDocument *destination;
            // Microsoft Office versions for Microsoft Windows OS
            uint vOffice2000 = 9, vOffice2002 = 10, // vOffice2003 = 11,
            vOffice2007 = 12, vOffice2013 = 15;
            
            // WdCompareTarget
            //var wdCompareTargetSelected = 0;
            //var wdCompareTargetCurrent = 1;
            WordWdCompareTarget wdCompareTargetNew = WordWdCompareTargetCompareTargetNew;
            // WdViewType
            WordWdViewType wdMasterView = WordWdViewTypeMasterView;
            WordWdViewType wdNormalView = WordWdViewTypeNormalView;
            WordWdViewType wdOutlineView = WordWdViewTypeOutlineView;
            // WdSaveOptions
            WordSaveOptions wdDoNotSaveChanges = WordSaveOptionsNo;
            //var wdPromptToSaveChanges = -2;
            //var wdSaveChanges = -1;
            WordWdViewType wdReadingView = WordWdViewTypeWordNoteView; //7;
            
            const char** objArgs = &argv[1];
            int num = argc - 1;
            if (num < 2)
            {
                NSString *basename = @(argv[0]); //[NSString stringWithUTF8String:argv[0]];
                printf("Usage: %s docdiff-osx base.doc new.doc\n", [[[basename lastPathComponent] stringByDeletingPathExtension] UTF8String]);
                return 1;
            }
            
            sBaseDoc = @(objArgs[0]);
            sNewDoc = @(objArgs[1]);
            
            
            if (![[NSFileManager defaultManager] fileExistsAtPath:sBaseDoc])
            {
                printf("File %s does not exist.  Cannot compare the documents.\n", [sBaseDoc UTF8String]);
                return 1;
            }
            
            if (![[NSFileManager defaultManager] fileExistsAtPath:sNewDoc])
            {
                printf("File %s does not exist.  Cannot compare the documents.\n", [sNewDoc UTF8String]);
                return 1;
            }
            
            @try
            {
                word = [SBApplication applicationWithBundleIdentifier:@"com.microsoft.Word"];
                
                if ([[word version] intValue] >= vOffice2013)
                {
                    if (![[NSFileManager defaultManager] isWritableFileAtPath:sBaseDoc])
                    {
                        // reset read-only attribute
                        [[NSFileManager defaultManager] setAttributes:@{NSFileImmutable: [NSNumber numberWithBool:NO]}ofItemAtPath:sBaseDoc error:nil];
                    }
                }
            }
            @catch(NSException * e)
            {
            }
            //@finally
            //{
            //}
            
            if ([[word version] intValue] >= vOffice2007)
            {
                sTempDoc = sNewDoc;
                sNewDoc = sBaseDoc;
                sBaseDoc = sTempDoc;
            }
            
            // The "visible" property does not exist in this interface
            //[word visible]
            
            // Open the new document
            @try
            {
                destination = [word open:nil fileName:sNewDoc confirmConversions:YES readOnly:([[word version] intValue] < vOffice2013) addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];
            }
            @catch(NSException * e)
            {
                
                @try
                {
                    // open empty document to prevent bug where first Open() call fails
                    destination = [word activeDocument];
                    destination = [word open:nil fileName:sNewDoc confirmConversions:YES readOnly:([[word version] intValue] < vOffice2013) addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];
                }
                @catch(NSException * e)
                {
                    printf("Error opening %s\n", [sNewDoc UTF8String]);
                    // Quit
                    return 1;
                }
            }
            
            // If the Type property returns either wdOutlineView or wdMasterView and the Count property returns zero, the current document is an outline.
            
            if ((([[[destination activeWindow] view] viewType] == wdOutlineView) || (([[[destination activeWindow] view] viewType] == wdMasterView) || ([[[destination activeWindow] view] viewType] == wdReadingView))) && ([[destination subdocuments] count] == 0))
            {
                // Change the Type property of the current document to normal
                [[[destination activeWindow] view] setViewType:wdNormalView];
            }
            
            // Compare to the base document
            if ([[word version] intValue] <= vOffice2000)
            {
                // Compare for Office 2000 and earlier
                @try
                {
                    // Contrary to the original TortoiseSVN/Git script, we cannot use duck typing -> comment out this line,
                    // as we only support the newer interface below
                    
                    //[destination comparePath:sBaseDoc];
                    
                }
                @catch(NSException * e)
                {
                    printf("Error comparing %s and %s\n", [sBaseDoc UTF8String], [sNewDoc UTF8String]);
                    // Quit
                    return 1;
                }
            }
            else
            {
                // Compare for Office XP (2002) and later
                @try
                {
                    [destination comparePath:sBaseDoc authorName:@"Comparison" target:wdCompareTargetNew detectFormatChanges:YES ignoreAllComparisonWarnings:YES addToRecentFiles:NO];
                }
                @catch(NSException * e)
                {
                    printf("Error comparing %s and %s\n", [sBaseDoc UTF8String], [sNewDoc UTF8String]);
                    // Close the first document and quit
                    [destination closeSaving:wdDoNotSaveChanges savingIn:nil];
                    return 1;
                }
            }
            
            // Show the comparison result
            if ([[word version] intValue] < vOffice2007)
            {
                [[[word activeDocument] windows][0] setVisible:YES];
            }
            
            // Mark the comparison document as saved to prevent the annoying
            // "Save as" dialog from appearing.
            [[word activeDocument] setSaved:YES];
            
            // Close the first document
            if ([[word version] intValue] >= vOffice2002)
            {
                [destination closeSaving:wdDoNotSaveChanges savingIn:nil];
            }
        }
    }
    return 0;
}
