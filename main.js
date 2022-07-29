/* 
Apps Script 로 동작하는 간단한 자동화 툴입니다.
Google Sheets 와 Google Docs 문서를 연계하여,
시트 데이터의 내용을 문서에 연동하고, 이를 pdf 로 출력할 수 있습니다.
*/

const ui = SpreadsheetApp.getUi(); // 버튼 제작을 위한
const folder = DriveApp.getFolderById(`[id 값을 넣어주세요.]`); // 작업을 저장할 폴더의 id값
const templateFileId = `[id 값을 넣어주세요.]`; // 템플릿 파일의 id값
const templateSheet =
	SpreadsheetApp.openById(`[id 값을 넣어주세요.]`).getSheetByName(
		`[시트 이름을 넣어주세요]`
	); // 변경값이 들어간 데이터 시트 정보
const lastRow = templateSheet.getLastRow(); // 데이터 시트의 마지막 행 값
const data = templateSheet.getRange(`A1:G${lastRow}`).getValues(); // 데이터 시트의 값을 이중 배열로 불러옵니다.
const head = data.shift(); // 데이터에서 제목을 분리합니다.

// 버튼과 연결된 메인 함수
function main() {
	let response = ui.alert("문서를 제작하시겠습니까?", ui.ButtonSet.YES_NO);
	if (response === ui.Button.YES) {
		createDocument();
		ui.alert(`문서 제작 완료`);
	} else {
		ui.alert(`제작을 취소합니다.`);
	}
}

// 문서 제작 함수
function createDocument() {
	let count = 0;
	try {
		data.forEach((row) => {
			// 문서 생성을 위해 템플릿 파일을 임시로 복제합니다.
			let documentId = DriveApp.getFileById(templateFileId).makeCopy().getId();
			let document = DocumentApp.openById(documentId);
			let body = document.getBody();

			// 복제한 파일의 placeholder를 데이터에 따라 변경합니다.
			// body.replaceText(`{{문서번호}}`, `${row[1]}`) 이런 식으로 반복
			row.forEach((col, idx) => {
				body.replaceText(head[idx], col);
			});

			DriveApp.getFileById(documentId).setName(`${row[2]}`); // 데이터 시트에 따라 변경 가능
			document.saveAndClose();

			// doc 파일을 pdf 로 변환하는 과정입니다.
			let source = DriveApp.getFileById(documentId);
			let blob = source.getAs(`application/pdf`);
			let newPDFFile = folder.createFile(blob);
			Logger.log(`${row[2]} 문서가 제작되었습니다.`);

			// 임시로 생성했던 doc 파일을 삭제합니다.
			DriveApp.getFileById(documentId).setTrashed(true);
			count++;
		});
	} catch (error) {
		Logger.log(error);
	}
	Logger.log(`총 ${count}개의 문서가 제작되었습니다.`);
	return;
}
