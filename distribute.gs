  function distribute(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet().toast('分配進行中，請稍等。', '狀態', -1);
    var sheet_A = ss.getSheetByName('報表A區');              //表單:“報表Ａ區”
    var sheet_V = ss.getSheetByName('Vendor');              //表單:"Vendor"

    var values_A = sheet_A.getDataRange().getValues();      //值:“報表Ａ區”
    var values_V = sheet_V.getDataRange().getValues();      //值:"Vendor"

    var item_A = values_A[0].indexOf('Item');               //品名:“報表Ａ區”
    var item_V = values_V[0].indexOf('Item');               //品名:"Vendor"

    var qty_A = values_A[0].indexOf('sum #VALUE!');         //數量:“報表Ａ區”
    var qty_V = values_V[0].indexOf('Open Quantity');       //數量:"Vendor"

    var vendor_note_V = values_V[0].indexOf('Vendor Note'); //1FS : "Vendor Note"
    var item_note_V = values_V[0].indexOf('Item Note');     //1FS  : "item Note"

    var comp_V = values_V[0].indexOf('comp');               //Vendor ：比對數量 
    var PO_V = values_V[0].indexOf('PO');                   //Vendor ：單號

    var iferror = []                                        //錯誤陣列
    var ValueSet = []                                       //正確陣列
    
    /*
      四個for
      第一個 : 由報表A區開始每一項比較

      第二個 : 處理1FS是否超過報表A區的數量
      第三個 : 處理[Vendor Note]有1FS字樣的分配
      第四個 : 處理有數量且無1FS字樣的分配
    */

    for(let a=1;a<values_A.length;a++)
    {
     if(values_A[a][item_A]!=''||values_A[a][item_A]!=0)
     {
      var qtyCompare = values_A[a][qty_A];
      var Cond = 0;
      //處理 [Vendor note] 有 1FS
      if (qtyCompare>0)
      {
        let totolQty = 0;


        //處理 1FS 是否超過 報表A區 的數量
        for(let v=0;v<values_V.length;v++){
            if((values_A[a][item_A] == values_V[v][item_V])&& 
               (values_V[v][vendor_note_V].toString().toUpperCase().indexOf('1FS')>-1)){

                 let upperValue = values_V[v][vendor_note_V].toString().toUpperCase();

                  if(upperValue.indexOf('TBA')>-1)
                  {
                    
                    let compareNum = trueQty(upperValue)
                    totolQty += compareNum;
                    if(values_A[a][qty_A]<compareNum || totolQty > values_A[a][qty_A]){

                        iferror.push(
                        `Vendor:[${v+1}]列開始 已填 '1FS' 的Vendor note,已超過
                         報表A區:[${a+1}]列,item:[${values_A[a][item_A]}]的數量\n`);

                         
                         break;
                        }     
                        else if(values_A[a][qty_A]>compareNum && values_A[a][qty_A]>totolQty){

                         totolQty += compareNum;
                        }
                  }
                  else if(upperValue == '1FS')
                  {
                    totolQty += values_V[v][qty_V];
                    if(values_A[a][qty_A]<values_V[v][qty_V] || totolQty > values_A[a][qty_A])
                    {
                        iferror.push(
                       `Vendor:[${v+1}]列開始 已填 '1FS' 的Vendor note,已超過
                        報表A區:[${a+1}]列,item:[${values_A[a][item_A]}]的數量\n`);
                        break;
                    }
                    else if(values_A[a][qty_A]>values_V[v][qty_V] && values_A[a][qty_A]>totolQty)
                    {

                        totolQty += values_V[v][qty_V];
                    }
                  }
              }  
        }


        //處理 [Vendor Note] 有1FS字樣 的分配
        for(let v =0;v<values_V.length;v++)
           {
           if(values_A[a][item_A] == values_V[v][item_V] &&
              values_V[v][vendor_note_V].toString().toUpperCase().indexOf('1FS')>-1)
                {
                  if(values_V[v][vendor_note_V].toString().toUpperCase() == '1FS')
                  {
                      if(qtyCompare<values_V[v][qty_V] )
                      {
                        Cond++ ;
                        break;
                      }
                      else if(qtyCompare>values_V[v][qty_V])
                      {
                        PushAndDecrease(v,1,values_V[v][vendor_note_V],0)
                      }
                  }
                  else if(values_V[v][vendor_note_V].toString().toUpperCase().indexOf('TBA')>-1)
                  {
                      if(qtyCompare<values_V[v][qty_V] )
                      {
                        Cond++ ;
                        break;
                      }
                      else if(qtyCompare>values_V[v][qty_V])
                      {
                        PushAndDecrease(v,1,values_V[v][vendor_note_V],1)
                      }
                  }
                }
           }
      }

      
      //處理 [item Note] 有1FS字樣 的分配
      if (Cond==0)
      {
        for(let v =0;v<values_V.length;v++)
           {
            if(values_A[a][item_A] == values_V[v][item_V]&&
               values_V[v][vendor_note_V] == '' &&             
               values_V[v][item_note_V].toUpperCase().toString().indexOf('1FS')>-1)
              {
                if(qtyCompare<values_V[v][qty_V] )
                {
                  Cond++ ;
                  break;
                }
                else if(qtyCompare>values_V[v][qty_V])
                {
                  PushAndDecrease(v,2,values_V[v][vendor_note_V],0)
                }
              }
            }
      }

     //處理有數量且 無 1FS字樣 的分配
      if(Cond==0)
      {
          for(let v=0;v<values_V.length;v++)
          {
              if(values_V[v][vendor_note_V] == ''&&
                 values_A[a][item_A] == values_V[v][item_V]&& 
                 values_V[v][item_note_V].toUpperCase().toString().indexOf("1FS")==-1)
              {
                  if(qtyCompare>values_V[v][qty_V])
                    PushAndDecrease(v,3,values_V[v][vendor_note_V],0);
                  else if(qtyCompare<values_V[v][qty_V])
                    break;
              }
          }  
      }

      


     } 

    }


      //處理setValue 
      function PushAndDecrease(v,p,pretext,tba){
          if(tba == 0){
              if(pretext!=''){
                 ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext}`,'comp':`${qtyCompare}`})
              }
              else if(pretext == ''){
                 ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`1FS`,'comp':`${qtyCompare}`})
              }
          qtyCompare-=values_V[v][qty_V];
          }
          else if(tba == 1){
              let string = pretext;
              let strMouse = string.indexOf('@')+1;
              let strSlice = string.indexOf('/');
              var newQty = parseInt(string.slice(strMouse,strSlice), 10);


              if(newQty<=values_V[v][qty_V])
              ValueSet.push({"row":v+1,"col":vendor_note_V+1,"value":`${pretext}`,'comp':`${qtyCompare}`})
              else if(newQty>values_V[v][qty_V])
              iferror.push([`\n第 ${v+1} 列，[Vendor Note]輸入數量有誤 \n`]);

              qtyCompare-=newQty;
          }
         
      }


      //iferror push的處理
      if(iferror.length>0)
      {
          SpreadsheetApp.getUi().alert(`${iferror} \n\n 程式中斷，請檢查後再執行!`);
          SpreadsheetApp.getActiveSpreadsheet().toast('錯誤，分配結束', '狀態', 5);
      }
      else if(iferror.length == 0)
      {
          sheet_V.getRange(2,comp_V+1,sheet_V.getLastRow(),1).clearContent();
        for(let i=0;i<ValueSet.length;i++)
        {
          let row = ValueSet[i].row;
          let col = ValueSet[i].col;
          let values = ValueSet[i].value
          let compQty = ValueSet[i].comp
          sheet_V.getRange(row,col).setValue(values);
          sheet_V.getRange(row,comp_V+1).setValue(compQty)
        }
        SpreadsheetApp.getUi().alert('分配已完成');
        SpreadsheetApp.getActiveSpreadsheet().toast('分配已完成', '狀態', 5);
      }   
  }


//起始alert
function distribute_Alert()
{
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert
    ('是否要進行自動分配?', 
     `自動分配的方式為
      先從[Vendor Note]確認是否有"1FS"字樣，
      若在該格有"1FS"，擇優先調配給該單號，
      若還有入量再接續給[Item Note]有"1FS"分配 。
      
      分配的方式由上到下，
      若上一個配完且數量小於下一單的
      [Open Quantity]的數量則中斷程序，
      由報表A區的下一個品項開始進行分配。\n
      
      確定要進行自動分配嗎?`, 
     ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (response == ui.Button.YES) 
    {
      distribute();
    } else 
    {
      SpreadsheetApp.getActiveSpreadsheet().toast('已取消執行', '狀態', 3);
    }
}

//處理TBA格式
function trueQty(value){
  let string = value;
  let strMouse = string.indexOf('@')+1;
  let strSlice = string.indexOf('/');
  let newQty = parseInt(string.slice(strMouse,strSlice), 10);
  return newQty;
}











