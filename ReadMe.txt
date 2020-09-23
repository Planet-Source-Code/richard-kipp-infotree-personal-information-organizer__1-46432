InfoTree - Personal Information Organizer   06/25/2003

Developer Info:
Written By:  Richard Kipp
Tools:       Visual Basic Pro 6.0(SP3)  (MS-Access Pro '97(SR2) was also loaded on the PC, but shouldn't be necessary for this App).  

This project was developed in an effort to understand LDAP tree structures.  This app is NOT LDAP compliant, but uses similar philosophies to attain high-speed data storage and retrieval.

The InfoTree lets you store and retrieve text data, and organize it as you wish. Features (many are color-coded) include:
  - Moving, copying, deleting (with or without all sub-branches) and sorting branches.
  - Exporting to Text File, Publishing as HTML List.
  - Add, Edit, Delete, Find, and Bold Text data.
  - Cross-linked (aliased) branches or nodes.
  - Password protected Nodes (hint: Use this on the parent of the node to be protected).
  - Automatic In-tree help and tree statistics.

Caveats:
you'll need to add the following components to use InfoTree in your project (already referenced in this demo project): 
1) Microsoft Windows Common Controls 6.0 (SP3)
   (MSCOMCTL.OCX)
2) Microsoft DAO 3.5 Object Library (or later)
   (DAO350.DLL)
The *Replace menu function has not yet been coded. After making this a 'stand-alone' application (it has been in use for nearly a year as part of a larger App), I noticed during testing that the print key actually sends the data you requested to InfoTree.txt on your desktop, then opens NotePad to view/print it.
