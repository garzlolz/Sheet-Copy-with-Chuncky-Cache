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
//     /*** 
//     [最外層的for]
//     從報表A開始 每一列做比對
//     定義一個剩下多少得值 qtyCompare為 當筆報表A區item的數量
//     定義Cond 當 Cond==0 表示數量充足 Cond == 1表示數量已經不足

//     [迴圈內第一個for]
//         Vendor表的
//         Vendor Note是否為空值 且
//         報表A區的item與Vendor的item 是否相符合 且
//         item_note 內是否已經有'1FS'若有 且
//         qtyCompare >=Vendor item 的數量
    
//         是的話優先發配給他
//         set Vendor Note 為'1FS'
//         每次發完後 qtyCompare 會減一次 

//         如果qtyCompare在下一筆少於Open QTY
//         則 Cond +1

//      [迴圈內第二個for]
    
//         與上面條件差在Vendor Note 非 1FS
//         且先判斷Cond 是否為0

//         是的話優先發配給他
//         set Vendor Note 為'1FS'
//         每次發完後 qtyCompare 會減一次  


//      ***/
//     for(let a=1;a<values_A.length;a++)
//     {
//      if(values_A[a][item_A]!=''||values_A[a][item_A]!=0)
//      {
//       var qtyCompare = values_A[a][qty_A];
//       var Cond = 0

//       var count = 0
//       //這個for 用來判斷是否重複
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

//       if(values_A[a][qty_A] == count && values_A[a][qty_A]!='' && values_A[a][item_A]!=''){
//         iferror.push(`第 ${a} 列 item ${values_A[a][item_A]}重複輸入`)
//       };
//       console.log(count);
//       console.log(values_A[a][qty_A])

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
//                  ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext},Compare = ${qtyCompare-values_V[v][qty_V]}, Priority:${p}`})
//               }
//               else if(pretext == ''){
//                  ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`1FS,Compare = ${qtyCompare-values_V[v][qty_V]}, Priority:${p}`})
//               }
//           qtyCompare-=values_V[v][qty_V];
//           }
//           else if(tba == 1){
//               let string = pretext;
//               let strMouse = string.indexOf('@')+1;
//               let strSlice = string.indexOf('/');
//               var newQty = parseInt(string.slice(strMouse,strSlice), 10);


//               if(newQty<=values_V[v][qty_V])
//               ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext},Compare = ${qtyCompare-newQty}, Priority:${p}`})
//               else if(newQty>values_V[v][qty_V])
//               iferror.push([`第 ${v+1} 列，[Vendor Note]輸入數量有誤 \n`]);

//               qtyCompare-=newQty;
//           }
         
//       }

//       if(iferror.length>0){
//           // SpreadsheetApp.getUi().alert(`${iferror} 程式中斷，請檢查後再執行!`);
//           // SpreadsheetApp.getActiveSpreadsheet().toast('錯誤，分配結束', '狀態', 5);
//       }
//       else if(iferror.length == 0){
//         for(let i=0;i<ValueSet.length;i++){
//             let row = ValueSet[i].row;
//             let col = ValueSet[i].col;
//             let values = ValueSet[i].value
//             sheet_V.getRange(row,col).setValue(values);
//         }
//         // SpreadsheetApp.getUi().alert('分配已完成');
//         // SpreadsheetApp.getActiveSpreadsheet().toast('分配已完成', '狀態', 5);
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

// // function ewqeqw(){
// //   // var ss = SpreadsheetApp.getActiveSpreadsheet();
// //   // var sheet = ss.getSheetByName('Vendor');
// //   // var values = sheet.getDataRange().getValues();
// //   // var note = values[0].indexOf('Vendor Note');

// //   // console.log(values[507][note])

// //   var string = '1FS@100/@1 TBA';
// //   var mouse = string.indexOf('@');
// //   var slice = string.indexOf('/');
// //   var newString = string.slice(mouse+1,slice); 
// //   console.log(newString);
// // }










