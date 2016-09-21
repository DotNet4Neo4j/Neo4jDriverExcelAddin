# Neo4j-Driver Excel Addin

This is a super early super alpha super use this as a starting point example of using the `Neo4j-Driver` nuget package 
within a VSTO (Visual Studio Tools for Office) application (in this case *Excel*).

This is mainly just a proof of concept at the moment.

## Things to note:

* Server location is hard coded (to `bolt://localhost`)
* The Column results are put into is hard coded (to `A`)
* The Column retrieved from the Cypher is hard coded (to `UserId`)

## Example

![alt text][examplegif]

## To Use

Clone the code, open in Visual Studio, and Run. This will start up Excel and then simply click on the 'Add-ins' menu in the 
ribbon. You will need to provide Cypher that returns a single column with the alias of `UserId` for example:

`MATCH (u:User) RETURN u.Id AS UserId LIMIT 10`

or

`MATCH (h:House) RETURN h.Address AS UserId`

This is due to the hard coding of the column retrieved.

# !!! REMEMBER - THIS IS A PROOF OF CONCEPT !!!


[examplegif]: https://p5iwog-ch3302.files.1drv.com/y3mIDySDI5TkjaBoPPstG_LLqDA94VNFbzV-nccm_3B70lu445RCrmZNacSeU9QVYy1Bpz-sw4qHXwKDojVLB38HjWtSdjwhLQr-PQZjjwpqTV6MezuyfmVNN51Yauy2KPT5Pn5uHQPq6grk2Jjycu6qCqOYLETkNJ6my1ddSrCyefkLeM_N7jzjYaFtDOGBNbJ/improvedExcelAddin.gif?psid=1 "Example of the addin working"