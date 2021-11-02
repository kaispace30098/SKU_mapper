# SKUmapper
Our company are using the main SKU and different SKUs for your selling channels which can help us manage a product that you sell on multiple online marketplaces.
This VB code saved in .txt file is dedicated to map sku valuse for each sales channel to be the SKU that is set for that product.
![alt text](https://cdn.shopify.com/app-store/listing_images/680ad538451da0ef1bb419cbe8b36254/desktop_screenshot/CIGC4c30lu8CEAE=.png?height=900&width=1600)
In this case, we export the weekly sales from our shipping software, ShipStation in excel with column headers (least requirement):

1.Quantity	
2.SKU	
3.StoreName

You also need to add two excel Sheets. and copy the original data and paste into 1st sheet
The first sheet after run the code will be the outcome for user to convert into a pivot table for integrated information.
The second sheet will be used only for data processing purpose.


The code contains four parts,

## Sub del() 
The function is used to delete the rows which are not the products but accessories and returns or refurbished items
For example, if product123Ring is the Ring of the product123 and product456_Pot is the Pot of product456.
We can put Ring and Pot into If statement in this section, where column "C" is the SKU column

        If InStr(Sheet2.Cells(i, "C"), "Ring") + InStr(Sheet2.Cells(i, "C"), "Pot") <> 0 Then
            Sheet2.Rows(i).EntireRow.Delete
        End If

## Sub SKUandQ()
This function is try to Unite the same items with same SKU, & multiply qty by any qty shown in the string of the sku
For example, if product product 'GW22921' from one Channel is named: "GW22921-S", and product 'F3528002' in other channels offers 6 quantity variants called "A-F3528002-06"
We can put them into the Select statement as following, where column "C" is the SKU column and column "B" is the SKU column
      
      Select Case Sheet2.Cells(i, "C")
            
            Case "GW22921"
                Sheet2.Cells(i, "C") = "GW22921-S"
                             
            Case "A-F3528002-06"
                Sheet2.Cells(i, "C") = "F3528002"
                Sheet2.Cells(i, "B") = Sheet2.Cells(i, "B") * 6 
        
      End Select
      
 ## Sub convert_bundle()
 This function is try to seperate the multiple SKUs with in a SKU used as bundle sales. The idea is to use the thrid sheet(Sheet2) to add line for the extra bundle item, and paste it back to Sheet2.
 First we loop over the second sheet Sheet1 and if in row i, we find out there is a bundle sku, for example:"N-GW22622+N-GWA0007", we remain the first sku" GW22622" in the Sheet1, and put the second item into the Sheet2 from row i.
 After the iteration, we delete the empty row in Sheet2 and from the last rows of Sheet1, we copy all data from Sheet2 and paste under it.
 
    LastRow = Sheet2.Cells(Rows.Count, "a").End(xlUp).Row 'find last row of Sheet1
      For i = LastRow To 2 Step -1 'loop thru backwards, finish at 2 for headers
    
          Select Case Sheet2.Cells(i, "C")
              Case "N-GW22622+N-GWA0007"
                  Sheet2.Cells(i, "C") = "GW22622"
                  Sheet3.Cells(i, "A") = Sheet2.Cells(i, "A")
                  Sheet3.Cells(i, "B") = Sheet2.Cells(i, "B")
                  Sheet3.Cells(i, "C") = "GWA0007"
                  Sheet3.Cells(i, "D") = Sheet2.Cells(i, "D")
                  Sheet3.Cells(i, "E") = Sheet2.Cells(i, "E")
                  Sheet3.Cells(i, "F") = Sheet2.Cells(i, "F")
          End Select
       Next i
    
    LR3 = Sheet3.Cells(Rows.Count, "A").End(xlUp).Row 'find last row of Sheet3
        
    For i = LR3 To 1 Step -1
        If Sheet3.Cells(i, "b") = "" Then
            Sheet3.Rows(i).EntireRow.Delete
        End If
    Next i
    
    LR3 = Sheet3.Cells(Rows.Count, "A").End(xlUp).Row 'find last row
    
    For k = 1 To LR3
        Sheet2.Cells(LastRow + k, "A") = Sheet3.Cells(k, "A")
        Sheet2.Cells(LastRow + k, "B") = Sheet3.Cells(k, "B")
        Sheet2.Cells(LastRow + k, "C") = Sheet3.Cells(k, "c")
        Sheet2.Cells(LastRow + k, "D") = Sheet3.Cells(k, "D")
        Sheet2.Cells(LastRow + k, "E") = Sheet3.Cells(k, "E")
        Sheet2.Cells(LastRow + k, "F") = Sheet3.Cells(k, "F")
    Next k
## Sub change_name()
If the company decide to change the name of the master sku we can use the function to modify it at last section.
For example, if company wants to change the master sku "GW44800-O' to "GW44800", we can put "GW44800" into the if statement
    
        Select Case Sheet2.Cells(i, "c")
            Case "GW44800-O"
                Sheet2.Cells(i, "c") = "GW44800"
        End Select
    Next i
    
## Sub callall()
Call all the functions we have discussed above!



 
