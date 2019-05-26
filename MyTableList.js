var tName = '';
const tSchema = 'test'; // 環境に合わせて変える
const tDatSrc ='Provider=MSDASQL; Data Source=Connector_MariaDB'; // 環境に合わせて変える
// テーブル一覧画面
function setList() {
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT TABLE_NAME,TABLE_COMMENT,TABLE_ROWS,DATE_FORMAT(CREATE_TIME,'%Y/%m/%d %H:%i:%s')"
            + " FROM information_schema.TABLES WHERE TABLE_SCHEMA = '" + tSchema + "'";
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    alert('対象テーブル検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    return;
  }
  if (rs.EOF){
    rs.Close();
    cn.Close();
    rs = null;
    cn = null;
    clrScr();
    $('#tabs').tabs( { active: 1} );
    return;
  }
  var strDoc = '';
  while (!rs.EOF){
    strDoc += '<tr><td style="width:150px;"><a href="#" onClick=colPage("' + rs(0).value + '")>' + rs(0).value + '</a></td>';
    strDoc += '<td width="300">' + rs(1).value + '</a></td>';
    strDoc += '<td width="80" align="RIGHT">' + rs(2).value + '</a></td>';
    strDoc += '<td width="200">' + rs(3).value + '</a></td></tr>';
    rs.MoveNext();
  }
  $('#lst01').replaceWith('<tbody id="lst01">' + strDoc + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  strDoc = '';
  $('#tabs').tabs( { active: 0} );
  $('#li02').css('visibility','hidden');
}
// テーブル項目詳細画面
function colPage(tName) {
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT COLUMN_COMMENT,COLUMN_NAME,COLUMN_TYPE"
            + " FROM information_schema.COLUMNS"
            + " WHERE TABLE_SCHEMA = '" + tSchema + "' AND TABLE_NAME = '" + tName + "'";
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  var strDoc = '';
  var strDoc1 = '';
  var strDoc2 = '';
  var strDoc3 = '';
  var cmtFlg = 0;
  while (!rs.EOF){
    if (rs(0).value != '') { cmtFlg = 1; }   // 項目コメントの有無
    var dtype = rs(2).value;
    var txtNum = 60;
    if (dtype == 'date') {
      txtNum = 90;
    } else if (dtype == 'time') {
      txtNum = 70;
    } else if (dtype == 'datetime') {
      txtNum = 130;
    } else if (dtype == 'text') {
      txtNum = 400;
    } else if (dtype.slice(0,7) == 'varchar') {
      txtNum = dtype.slice(8,(dtype.length -1)) * 8 + 10;
      if (txtNum > 400) { txtNum = 400; }
      if (txtNum < 80) { txtNum = 80; }
    }
    strDoc1 += '<td style="width:' + txtNum + 'px;">' + rs(0).value + '</td>';
    strDoc2 += '<td style="width:' + txtNum + 'px;">' + rs(1).value + '</td>';
    strDoc3 += '<td nowrap>' + dtype + '</td>';
    rs.MoveNext();
  }
  if (cmtFlg == 0) {
    strDoc = '<tr style="display: none;">' + strDoc1 + '<td class="dummyColumn"></td></tr>'
  } else {
    strDoc = '<tr class="bg-primary">' + strDoc1 + '<td class="dummyColumn"></td></tr>'
  }
  strDoc += '<tr class="bg-primary">' + strDoc2 + '<td class="dummyColumn"></td></tr>'
  strDoc += '<tr class="bg-primary">' + strDoc3 + '<td class="dummyColumn"></td></tr>';
  $('#hdr02').replaceWith('<tbody id="hdr02" style="color:white;">' + strDoc + '</tbody>');
  strDoc = '';
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  mySql = "SELECT * FROM " + tName;
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  while (!rs.EOF){
    strDoc += '<tr>';
    for ( var i = 0; i < rs.Fields.Count; i++ ) {
      if (rs(i).Type == 133) {
        strDoc += '<td style="width:90px;">' + formatDate(rs(i).value,'YYYY-MM-DD') + '</td>';
      } else if (rs(i).Type == 134) {
        strDoc += '<td style="width:70px;">' + formatDate(rs(i).value,'hh:mm') + '</td>';
      } else if (rs(i).Type == 135) {
        strDoc += '<td style="width:130px;">' + formatDate(rs(i).value,'YYYY-MM-DD hh:mm') + '</td>';
      } else if (rs(i).Type == 203) {
        strDoc += '<td style="width:400px;">' + rs(i).value + '</td>';
      } else if (rs(i).Type == 202) {
        var txtNum = rs(i).DefinedSize * 8 + 10;
        if (txtNum > 400) { txtNum = 400; }
        if (txtNum < 80) { txtNum = 80; }
        strDoc += '<td style="width:' + txtNum + 'px;">' + rs(i).value + '</td>';
      } else {
        strDoc += '<td style="width:60px;">' + rs(i).value + '</td>';
      }
    }
    rs.MoveNext();
  }
  strDoc += '</tr>';
  $('#lst02').replaceWith('<tbody id="lst02">' + strDoc + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  $('#tabs').tabs( { active: 1} );
  $('#li02').css('visibility','visible');
}
// 日付時刻のフォーマット
function formatDate(date, format) {
  var day = new Date(date);
  format = format.replace(/YYYY/, day.getFullYear());
  format = format.replace(/MM/, ('0' + (day.getMonth() + 1)).slice(-2));
  format = format.replace(/DD/, ('0' + day.getDate()).slice(-2));
  format = format.replace(/hh/, ('0' + day.getHours()).slice(-2));
  format = format.replace(/mm/, ('0' + day.getMinutes()).slice(-2));
  return format;
}
