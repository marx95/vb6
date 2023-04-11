MMM
===

"Make My Manifest"                              June 2006

A program that analyzes the project (*.vbp) file of a
compiled VB 6.0 EXE and produces a Registration Free COM
XCopy-deployable execution package from it.

                    -------------------------------------
              Note: Reg-Free COM only works on Windows XP
                    and later operating systems such as
                    Windows 2003 and Windows Vista.
                    -------------------------------------

An XCopy folder is created inside the project folder.
The EXE and all project ActiveX OCXs and DLLs are copied
into the folder.  An application manifest is created for
the EXE (which also enables XP Styles) along with an
assembly manifest for the ActiveX libraries.

Programs must still properly call the InitCommonControls
API in many cases to enable XP Styles.

Use of licensed ActiveX components may require manually
adding licenses within the EXE via the VB Licenses
collection.


Legal
=====

This program is free for all to use or modify.  It is
"as is" software, with no guarantee that it will not
result in damage to computers or associated data files.
No promise of support is implied and none should be
inferred.

This should be considered experimental software, not
yet complete and not successful in all scenarios.  Some
programs or component libraries are too complex for MMM
to process.  In particular MMM should not be used on
programs that make use of out of process ActiveX
components.


What's it For?
==============

MMM creates an execution package that can be deployed to
other Windows XP (or later) computers that have the VB
6.0 runtime components installed.  No "setup" or
installer program is required, nor is registration of
ActiveX components used by your program.

This makes it possible in many cases to run your VB 6.0
programs from media such as flash memory drives or CD-Rs.
This is done without needing to write anything to the
system Registry of target computers.


How Does it Work?
=================

MMM makes use of technologies Microsoft introduced with
Windows XP called Registration Free COM, Isolated
Applications, and Side by Side Assemblies.  While the
.Net editions of Visual Studio support these technologies
directly, features were never added to earlier versions
of the Visual Studio development suite such as Visual
Studio 6.0 to make it easy to use with older compilers.

It works perfectly well with many Win32 programs such as
VB programs however.  One simply needs to stay within
the bounds of these technologies' limitations.

The ActiveX libraries are gathered up as a "dependent
assembly" and a special XML file called an "assembly
manifest" is added.  Another XML file called an
"application manifest" is created for the EXE that
references the assembly manifest.

These manifests provide information about the
application and component libraries that would otherwise
need to be stored in the system Registry.


Operation
=========

Running MMM results in a file open dialog requesting a
VB 6.0 project (*.vbp) file for a compiled Standard EXE
project.  Several assumptions are made about the project:

    o The project has been successfully compiled to an
      EXE that is in the same folder as the project
      file.

    o Any ActiveX OCXs or DLLs the program uses are
      compiled, properly registered, and present on the
      computer running MMM.

    o Any licenses needed for controls are loaded into
      the VB Licenses collection at runtime by logic
      embedded in your program.

    o The XCopy-deployable execution package is to be
      created in a folder "XCopy" within the project
      folder.

    o Any data files, MDBs, conventional DLLs, or other
      files needed by the EXE at runtime will be copied
      into the "XCopy" folder manually after MMM is
      finished running.

    o Any special techniques required to open data files
      "read only" are handled by logic embedded in your
      program if your execution target is a read only
      medium such as CD-R.

After specifying the VB project to process MMM runs
to completion largely without further user interaction.
If the "XCopy" folder already exists you will be prompted
whether you want it overwritten.  If not, MMM will stop.
When complete MMM leaves a log of its analysis and
actions on screen until you close it.
