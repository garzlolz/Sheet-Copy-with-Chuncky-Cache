//   var property = PropertiesService.getDocumentProperties()
  
//   function distribute(){

//     var ss = SpreadsheetApp.getActiveSpreadsheet();
//     SpreadsheetApp.getActiveSpreadsheet().toast('分配進行中，請稍等。', '狀態', -1);
//     var sheet_A = ss.getSheetByName('報表A區');
//     var sheet_V = ss.getSheetByName('Vendor');

//     var values_A = sheet_A.getDataRange().getValues();
//     var values_V = sheet_V.getDataRange().getValues();

//     var item_A = values_A[0].indexOf('Item');
//     var item_V = values_V[0].indexOf('Item');

//     var qty_A = values_A[0].indexOf('sum #VALUE!');
//     var qty_V = values_V[0].indexOf('Open Quantity');

//     var vendor_note_V = values_V[0].indexOf('Vendor Note');
//     var item_note_V = values_V[0].indexOf('Item Note');
    
//     var iferror = []
//     var ValueSet = []


//     for(let a=1;a<values_A.length;a++)
//     {

//      if(values_A[a][item_A]!=''||values_A[a][item_A]!=0)
//      {

//       var qtyCompare = values_A[a][qty_A];

//       var Cond = 0
//       var count = 0;
//       //判斷是否重複
//       for(let v=0;v<values_V.length;v++ )
//       {
//         if(values_A[a][item_A] == values_V[v][item_V] &&
//            values_V[v][vendor_note_V].toString().toUpperCase().indexOf('1FS')>-1)
//         {

//             if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA')==-1)
//             {

//               count +=values_V[v][qty_V];
//             }
//             else if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA'))
//             {

//               let string = values_V[v][vendor_note_V].toString();
//               let mouse = string.indexOf('@')+1;
//               let slice = string.indexOf('/')+1;
//               let num = parseInt(string.slice(mouse,slice),10);
//               count+=num;
//             }
//         }
//       }
//       let date = new Date();
//       let dd = date.getDay();
//       let mm = date.getMonth();

//       // let obj = JSON.stringify({'item':values_A[a][item_A],'comp':count,'date':`${mm}/${dd}`});

//       var itemProperty = property.getProperty(values_A[a][item_A]);

      
//       if(itemProperty!=undefined)
//       {
//           itemProperty = JSON.parse(itemProperty);

//           console.log('itemProperty != undefined')
//           console.log(itemProperty);
//           console.log(typeof(itemProperty));
//           console.log('_____________________')

//         if(itemProperty.item == values_A[a][item_A] && itemProperty.comp == count && itemProperty.date == `${mm}/${dd}`){

//           console.log(itemProperty);
//           console.log(`itemProperty.comp == count=${itemProperty.comp == count}`)
//           console.log('_____________________')

//           iferror.push(`第 ${a} 列 ${values_A[a][item_A]} 重複輸入`);
//         }
//         // else if((itemProperty.item == values_A[a][item_A] && itemProperty.comp != count)){

//         //   console.log(`itemProperty.comp = ${itemProperty.comp}`);
//         //   console.log(`count:${count}`)
//         //   console.log('itemProperty.comp != count')
//         //   console.log('_____________________');

//         //   property.setProperty(values_A[a][item_A],obj);
//         // }
//       } 
//       //判斷重複結束


//       //處理 [Vendor note] 有 1FS
//       if (qtyCompare>0)
//       {

//         for(let v =0;v<values_V.length;v++)
//            {

//            if(values_A[a][item_A] == values_V[v][item_V] &&
//               values_V[v][vendor_note_V].toString().toUpperCase().indexOf('1FS')>-1)
//                 {

//                   if(values_V[v][vendor_note_V].toString().toUpperCase() == '1FS')
//                   {

//                       if(qtyCompare<values_V[v][qty_V] )
//                       {

//                         Cond++ ;
//                         break;
//                       }
//                       else if(qtyCompare>values_V[v][qty_V])
//                       {

//                         PushAndDecrease(v,1,values_V[v][vendor_note_V],0)
//                       }
//                   }
//                   else if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA')>-1)
//                   {

//                       if(qtyCompare<values_V[v][qty_V] )
//                       {

//                         Cond++ ;
//                         break;
//                       }
//                       else if(qtyCompare>values_V[v][qty_V])
//                       {

//                         PushAndDecrease(v,1,values_V[v][vendor_note_V],1)
//                       }
//                   }
//                 }
//            }
//       }

//       //處理 item Note 有 1FS
//       if (Cond==0)
//       {
//         for(let v =0;v<values_V.length;v++)
//            {
//             if(values_A[a][item_A] == values_V[v][item_V]&&
//                values_V[v][vendor_note_V] == '' &&             
//                values_V[v][item_note_V].toUpperCase().toString().indexOf('1FS')>-1)
//               {
//                 if(qtyCompare<values_V[v][qty_V] )
//                 {
//                   Cond++ ;
//                   break;
//                 }
//                 else if(qtyCompare>values_V[v][qty_V])
//                 {
//                   PushAndDecrease(v,2,values_V[v][vendor_note_V],0)
//                 }
//               }
//             }
//       }

//      //有數量且為空值
//       if(Cond==0)
//       {
//           for(let v=0;v<values_V.length;v++)
//           {
//               if(values_V[v][vendor_note_V] == ''&&
//                  values_A[a][item_A] == values_V[v][item_V]&& 
//                  values_V[v][item_note_V].toUpperCase().toString().indexOf("1FS")==-1)
//               {
//                   if(qtyCompare>values_V[v][qty_V])
//                     PushAndDecrease(v,3,values_V[v][vendor_note_V],0);
//                   else if(qtyCompare<values_V[v][qty_V])
//                     break;
//               }
//           }  
//       }

      



//      } 

//     }


//       //處理setValue
//       function PushAndDecrease(v,p,pretext,tba){
//           if(tba == 0){
//               if(pretext!=''){
//                 //  ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext},Compare = ${qtyCompare-values_V[v][qty_V]}, Priority:${p}`})
//                 ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext}`})
//               }
//               else if(pretext == ''){
//                 //  ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`1FS,Compare = ${qtyCompare-values_V[v][qty_V]}, Priority:${p}`})
//                 ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`1FS`})
//               }
//           qtyCompare-=values_V[v][qty_V];
//           }
//           else if(tba == 1){
//               let string = pretext;
//               let strMouse = string.indexOf('@')+1;
//               let strSlice = string.indexOf('/');
//               var newQty = parseInt(string.slice(strMouse,strSlice), 10);


//               if(newQty<=values_V[v][qty_V])
//               // ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext},Compare = ${qtyCompare-newQty}, Priority:${p}`})
//               ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext}`})
//               else if(newQty>values_V[v][qty_V])
//               iferror.push([`第 ${v+1} 列，[Vendor Note]輸入數量有誤 \n`]);

//               qtyCompare-=newQty;
//           }
         
//       }

//       if(iferror.length>0){

//           SpreadsheetApp.getUi().alert(`${iferror} 程式中斷，請檢查後再執行!`);
//           SpreadsheetApp.getActiveSpreadsheet().toast('錯誤，分配結束', '狀態', 5);
//       }
//       else if(iferror.length == 0){
//         for(let i=0;i<ValueSet.length;i++){
//             let row = ValueSet[i].row;
//             let col = ValueSet[i].col;
//             let values = ValueSet[i].value
//             sheet_V.getRange(row,col).setValue(values);
//         }

//         for(let a=1;a<values_A.length;a++)
//         {
//           var count = 0
//           if(values_A[a][item_A]!=''||values_A[a][item_A]!=0)
//           {
//             for(let v=0;v<values_V.length;v++ )
//             {
//               if(values_A[a][item_A] == values_V[v][item_V] &&
//                  values_V[v][vendor_note_V].toString().toUpperCase().indexOf('1FS')>-1)
//               {
              
//                   if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA')==-1)
//                   {
                  
//                     count +=values_V[v][qty_V];
//                   }
//                   else if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA')>-1)
//                   {
                  
//                     let string = values_V[v][vendor_note_V].toString();
//                     let mouse = string.indexOf('@')+1;
//                     let slice = string.indexOf('/')+1;
//                     let num = parseInt(string.slice(mouse,slice),10);
//                     count+=num;
//                   }
//               }
//             }
//            let date = new Date();
//            let dd = date.getDay();
//            let mm = date.getMonth();
  
//            let obj = JSON.stringify({'item':values_A[a][item_A],'comp':count,'date':`${mm}/${dd}`});
  
//            var itemProperty = property.getProperty(values_A[a][item_A]);
  
  
  
//            if(itemProperty == undefined){      
//              property.setProperty(values_A[a][item_A],obj);
//            }
//            else if(itemProperty!=undefined)
//            {
//              itemProperty = JSON.parse(itemProperty);
  
//              if(itemProperty.item == values_A[a][item_A] && 
//                 itemProperty.comp == count && 
//                 itemProperty.date == `${mm}/${dd}`){

//                 iferror.push(`第 ${a} 列 ${values_A[a][item_A]} 重複輸入`);
//              }
//              else if((itemProperty.item == values_A[a][item_A] && itemProperty.comp != count)||
//                      (itemProperty.item == values_A[a][item_A] && 
//                      itemProperty.comp == count && 
//                      itemProperty.date != `${mm}/${dd}`)){

//                 property.setProperty(values_A[a][item_A],obj);
//              }
             
//            }

//           }
//         }
//         SpreadsheetApp.getUi().alert('分配已完成');
//         SpreadsheetApp.getActiveSpreadsheet().toast('分配已完成', '狀態', 5);
//       }   
//   }




// function distribute_Alert(){

//     var ui = SpreadsheetApp.getUi();
//     var response = ui.alert
//     ('是否要進行自動分配?', 
//      `自動分配的方式為
//       先從[item Note]確認是否有"1FS"字樣，
//       若在該格有"1FS"，擇優先將報表A區的item數量調配給該單號，
//       若沒有則先等1FS先配完還有數量再依序分配。
      
//       分配的方式為由上到下，
//       若上一個配完的且數量小於下一單的
//       [Open Quantity]的數量則進行中斷，
//       由報表A區的下一個品項開始進行分配。
//       ⚠️ 若在Vendor Note有查到'1FS'字樣 最先列入計算\n
//       確定要進行自動分配嗎?`, 
//      ui.ButtonSet.YES_NO);

//     // Process the user's response.
//     if (response == ui.Button.YES) {
//       distribute();
//     } else {
//       SpreadsheetApp.getActiveSpreadsheet().toast('已取消執行', '狀態', 3);
//     }
// }


// function cp(){   
//   property.deleteAllProperties();
// }







