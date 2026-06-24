/****************************************************************
 * 그린맵 GreenMap — Google Apps Script 백엔드 (참고/붙여넣기용)
 *
 * 이 파일은 깃 저장소에 "참고용"으로 보관하는 문서입니다.
 * 실제로는 구글 스프레드시트의 [확장 프로그램 > Apps Script]에
 * 아래 내용을 붙여넣고 "배포 > 새 배포 > 웹 앱"으로 배포해야 합니다.
 *
 * 역할
 *   - doPost: form.html이 보낸 관찰기록 + 사진(base64)을 받아
 *             사진은 구글 드라이브에 저장(전체공개 링크)하고,
 *             시트에는 사진의 "URL"만 기록합니다.
 *   - doGet : 시트를 읽어 index.html이 쓰는 JSON(사진 URL 포함)으로 반환합니다.
 *
 * 시트 열 순서 (1행은 머리글, 아래 순서대로 만들어 두세요)
 *   A:타임스탬프 B:차시 C:관찰자 D:날짜 E:위치 F:종명 G:학명
 *   H:관찰내용 I:위도 J:경도 K:사진1 L:사진2 M:사진3
 ****************************************************************/

// ===== 설정 =====
// 1) 스프레드시트 ID (시트 주소 .../d/【여기】/edit 부분)
var SHEET_ID = '1ETBlzzmzEYOOYja9P6BvIuiIbHZfCeYezwPGseCZMwc';
var SHEET_NAME = '시트1'; // 데이터가 들어갈 시트 탭 이름
// 2) 사진을 저장할 드라이브 폴더 이름 (없으면 자동 생성)
var PHOTO_FOLDER = '그린맵_사진';

// ===== 권한 승인용 함수 (최초 1회만 실행) =====
// 사진 저장에는 구글 드라이브 권한이 필요합니다.
// Apps Script 편집기에서 함수 선택 → "권한승인" 선택 → ▶ 실행 →
// 팝업에서 본인 계정으로 드라이브/시트 접근을 "허용" 하세요.
// (이 과정을 거치지 않으면 학생 제출 시 'DriveApp 권한 없음' 오류가 납니다.)
function 권한승인() {
  // 드라이브 '수정(쓰기)' 권한까지 확실히 요청하려면 실제 쓰기 동작을 호출해야 합니다.
  // (읽기 동작만 호출하면 읽기 권한만 승인되어 사진 저장 시 createFolder 오류가 납니다.)
  var f = DriveApp.createFolder('그린맵_권한확인_삭제예정');
  f.setTrashed(true);                          // 테스트용 폴더는 휴지통으로
  SpreadsheetApp.getActiveSpreadsheet().getName(); // 시트 권한 확인/요청
  Logger.log('권한 승인 완료 — 이제 학생 제출이 정상 동작합니다.');
}

// ===== 폼 제출 받기 (form.html → 여기) =====
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    // 사진 base64 → 드라이브 저장 → 전체공개 URL
    var folder = getPhotoFolder_();
    var urls = ['', '', ''];
    var photos = data.photos || [];
    for (var i = 0; i < photos.length && i < 3; i++) {
      var p = photos[i];
      if (!p || !p.data) continue;
      var bytes = Utilities.base64Decode(p.data);
      var blob = Utilities.newBlob(bytes, 'image/jpeg', p.name || ('photo_' + (i + 1) + '.jpg'));
      var file = folder.createFile(blob);
      // "링크가 있는 모든 사용자" 보기 권한 → 지도에서 이미지 표시 가능
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls[i] = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    }

    sheet.appendRow([
      new Date(),
      data.session || '',
      data.observer || '',
      data.date || '',
      data.location || '',
      data.species || '',
      data.scientificName || '',
      data.observation || '',
      data.lat || '',
      data.lng || '',
      urls[0], urls[1], urls[2]
    ]);

    return json_({ status: 'ok' });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  }
}

// ===== 지도/목록에 데이터 주기 (index.html → 여기) =====
function doGet() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    var values = sheet.getDataRange().getValues();
    var out = [];
    for (var r = 1; r < values.length; r++) { // 0행은 머리글이므로 건너뜀
      var row = values[r];
      if (!row[5] && !row[8]) continue; // 종명·위도 둘 다 없으면 빈 행으로 보고 skip
      out.push({
        session: row[1],
        observer: row[2],
        date: formatDate_(row[3]),
        location: row[4],
        species: row[5],
        scientificName: row[6],
        observation: row[7],
        lat: row[8],
        lng: row[9],
        // index.html의 extractPhotos가 photos 배열을 그대로 읽습니다
        photos: [row[10], row[11], row[12]].filter(function(u) { return u; })
      });
    }
    return json_({ status: 'ok', data: out });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  }
}

// ===== 도우미 =====
function getPhotoFolder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function formatDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  return v || '';
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
