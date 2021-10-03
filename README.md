# word
A Pythonic Microsoft Word (.docx) file manager and style editor.

# Repository Details

* STATUS: Active
* LATEST RELEASE: v.1.0.0-beta
* LAST UPDATED: 2021-10-03
* LICENSE: Public Domain (except where otherwise noted)
* URL: https://github.com/dt-woods/word

# Summary
The goals of this project are to produce the following utility functions to automate some of the boring stuff with MS Word (.docx) files:

- [x] Read a .docx file with generic (built-in) styles and write a copy with custom styles preserving the content (i.e., map one style to another).
- [x] Read a .docx file and write individual .docx files based on a user-defined breaking style (e.g., parse a book into chapter files).
- [x] Read a list of .docx files and write into a single concatenated .docx file (e.g., merge chapters into a book).

A lower priority is to create supplemental utility functions that:

- [ ] Searches the content of one or more .docx files for abbreviations and save the results to a table (i.e., create a vocab table).

See separate issues for more details.

This project makes use of the `python-docx`, a Python API for Microsoft Word .docx files.
This [documentation][pydocx-doc] and [repository][pydocx-rep] are licensed and copyrighted by [Scanny][scanny] using the [MIT license][pydocx-lic].

[pydocx-doc]: https://python-docx.readthedocs.io/en/latest/#
[pydocx-lic]: https://github.com/python-openxml/python-docx/blob/master/LICENSE
[pydocx-rep]: https://github.com/python-openxml/python-docx
[scanny]: https://github.com/scanny
