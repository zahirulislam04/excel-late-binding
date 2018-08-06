## User often experience excel version dependency with Excel object used in source code because excel version varies in client machines. This can be resolved by late binding the the excel object in source code. My sample code shows late binding example written for data export from access form to excel.

Please note in my code snippents,

 - sfrmCandidateSearch is the access form name from where data will be exported to excel. So please replace this in your source code with appropriate.

 - following two lines used to copy data from access form, so you will need to write your own to replace it

	DoCmd.RunCommand acCmdSelectAllRecords
    	DoCmd.RunCommand acCmdCopy