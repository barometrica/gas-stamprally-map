// 使うデータの型だけ定義
type StampRallyResult = {
  properties: {
    Interviewer: {
      people?: {
        id: string;
      }[];
    };
    Interviewee: {
      people?: {
        id: string;
      }[];
    };
    Date: {
      date?: {
        start: string;
      };
    };
  };
};

type MemberResult = {
  properties: {
    アカウント: {
      people?: {
        id: string;
      }[];
    };
    Name: {
      title?: {
        plain_text: string;
      }[];
    };
    社員番号: {
      number: number;
    };
    旧メンバー: {
      checkbox: boolean;
    };
  };
  interviewerCount: number;
  intervieweeCount: number;
};

type StampRallyData = {
  results: StampRallyResult[];
  has_more: boolean;
  next_cursor: string | null;
};

type MemberData = {
  results: MemberResult[];
  has_more: boolean;
  next_cursor: string | null;
};

const START_ROW = 4;
const START_COLUMN = 4;

function getColName(num: number) {
  let sheet = SpreadsheetApp.getActiveSheet();
  return sheet.getRange(1, num).getA1Notation().replace(/\d/, "");
}

// スタンプラリーデータを取得
const fetchStampRallyData = (option?: { filter: object }) => {
  const token = PropertiesService.getScriptProperties().getProperty(
    "TOKEN_1ON1_STAMPRALLY_MAP"
  );
  const databaseId = PropertiesService.getScriptProperties().getProperty(
    "DATABASE_ID_1ON1_STAMPRALLY"
  );
  const url = `https://api.notion.com/v1/databases/${databaseId}/query`;

  const options = {
    headers: {
      Authorization: `Bearer ${token}`,
      "Notion-Version": "2022-06-28",
    },
    contentType: "application/json",
    method: "post" as const,
  };

  let results: StampRallyData["results"] = [];
  let nextCursor: StampRallyData["next_cursor"] = null;
  do {
    const data = JSON.parse(
      UrlFetchApp.fetch(url, {
        ...options,
        payload: JSON.stringify({
          filter: option?.filter,
          start_cursor: nextCursor ?? undefined,
        }),
      }).getContentText()
    ) as StampRallyData;
    results = [...results, ...data.results];
    nextCursor = data.next_cursor;
  } while (nextCursor);

  return results;
};

// メンバーデータを取得
const fetchMemberData = (option?: { filter: object }) => {
  const token = PropertiesService.getScriptProperties().getProperty(
    "TOKEN_1ON1_STAMPRALLY_MAP"
  );
  const databaseId =
    PropertiesService.getScriptProperties().getProperty("DATABASE_ID_MEMBER");
  const url = `https://api.notion.com/v1/databases/${databaseId}/query`;

  const options = {
    headers: {
      Authorization: `Bearer ${token}`,
      "Notion-Version": "2022-06-28",
    },
    contentType: "application/json",
    method: "post" as const,
  };

  let results: MemberData["results"] = [];
  let nextCursor: MemberData["next_cursor"] = null;
  do {
    const data = JSON.parse(
      UrlFetchApp.fetch(url, {
        ...options,
        payload: JSON.stringify({
          filter: option?.filter,
          start_cursor: nextCursor ?? undefined,
        }),
      }).getContentText()
    ) as MemberData;
    results = [...results, ...data.results];
    nextCursor = data.next_cursor;
  } while (nextCursor);

  return results
    .filter((result) => {
      return (
        !result.properties.旧メンバー.checkbox &&
        result.properties.Name.title?.at(0)?.plain_text &&
        result.properties.アカウント.people?.at(0)?.id
      );
    })
    .sort((a, b) => {
      if (
        (a.properties.社員番号.number === null ||
          a.properties.社員番号.number === undefined) &&
        (b.properties.社員番号.number === null ||
          b.properties.社員番号.number === undefined)
      ) {
        return 0;
      }
      if (
        a.properties.社員番号.number === null ||
        a.properties.社員番号.number === undefined
      ) {
        return 1;
      }
      if (
        b.properties.社員番号.number === null ||
        b.properties.社員番号.number === undefined
      ) {
        return -1;
      }
      return a.properties.社員番号.number - b.properties.社員番号.number;
    });
};

const init = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 出力対象のシートでなければ終了
  if (sheet.getName() !== "星取表") {
    const ui = SpreadsheetApp.getUi();
    ui.alert("シート「星取表」で実行してください");
    return;
  }

  const stampRallyData = fetchStampRallyData();
  const memberData = fetchMemberData();

  memberData.forEach((data) => {
    data.interviewerCount = 0;
    data.intervieweeCount = 0;
  });

  // スタンプラリーの組み合わせ配列を作成
  const stampRallyArray = stampRallyData.map((data) => {
    const interviewerId = data.properties.Interviewer.people?.at(0)?.id;
    const intervieweeId = data.properties.Interviewee.people?.at(0)?.id;
    const isDone =
      data.properties.Date.date?.start !== null &&
      data.properties.Date.date?.start !== undefined;
    return {
      interviewerId: interviewerId,
      intervieweeId: intervieweeId,
      isDone: isDone,
    };
  });

  // シートをクリア
  sheet.clearContents();

  // 星取表の配列を作成
  const memberLength = memberData.length;
  for (let row = 0; row < memberLength; row++) {
    const rowMember = memberData[row];
    for (let column = 0; column < memberLength; column++) {
      const columnMember = memberData[column];
      if (row === column) {
        sheet
          .getRange(
            `${getColName(column + START_COLUMN + 1)}${row + START_ROW + 2}`
          )
          .setValue("-");
      }
      const targetInterviewerId =
        rowMember.properties.アカウント.people?.at(0)?.id;
      const targetIntervieweeId =
        columnMember.properties.アカウント.people?.at(0)?.id;
      // 該当するスタンプラリーがあり実施済み
      if (
        stampRallyArray.some(
          (stampRally) =>
            stampRally.interviewerId === targetInterviewerId &&
            stampRally.intervieweeId === targetIntervieweeId &&
            stampRally.isDone
        )
      ) {
        memberData[row].interviewerCount = memberData[row].interviewerCount + 1;
        memberData[column].intervieweeCount =
          memberData[column].intervieweeCount + 1;
        sheet
          .getRange(
            `${getColName(column + START_COLUMN + 1)}${row + START_ROW + 2}`
          )
          .setValue("◎");
      }
      // 該当するスタンプラリーがあり実施済み
      if (
        stampRallyArray.some(
          (stampRally) =>
            stampRally.interviewerId === targetInterviewerId &&
            stampRally.intervieweeId === targetIntervieweeId &&
            !stampRally.isDone
        )
      ) {
        sheet
          .getRange(
            `${getColName(column + START_COLUMN + 1)}${row + START_ROW + 2}`
          )
          .setValue("予");
      }
    }
  }

  // 行のメンバー一覧を描画
  for (let row = 0; row < memberLength; row++) {
    const rowMember = memberData[row];
    sheet
      .getRange(`${getColName(START_COLUMN)}${row + START_ROW + 2}`)
      .setValue(rowMember.properties.Name.title?.at(0)?.plain_text);
    sheet
      .getRange(`${getColName(START_COLUMN - 1)}${row + START_ROW + 2}`)
      .setValue(rowMember.properties.社員番号.number);
    sheet
      .getRange(`${getColName(START_COLUMN - 2)}${row + START_ROW + 2}`)
      .setValue(rowMember.interviewerCount);
  }
  // 列のメンバー一覧を描画
  for (let column = 0; column < memberLength; column++) {
    const columnMember = memberData[column];
    sheet
      .getRange(`${getColName(column + START_COLUMN + 1)}${START_ROW}`)
      .setValue(columnMember.properties.Name.title?.at(0)?.plain_text);
    sheet
      .getRange(`${getColName(column + START_COLUMN + 1)}${START_ROW - 1}`)
      .setValue(columnMember.properties.社員番号.number);
    sheet
      .getRange(`${getColName(column + START_COLUMN + 1)}${START_ROW - 2}`)
      .setValue(columnMember.intervieweeCount);
  }

  const date = new Date();
  sheet
    .getRange(`A1`)
    .setValue(
      "更新日時：" +
        Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
    );
  sheet.getRange(`D3`).setValue("社員番号");
  sheet.getRange(`C4`).setValue("社員番号");
  sheet.getRange(`B4`).setValue("もらったスタンプの数");
  sheet.getRange(`D2`).setValue("あげたスタンプの数");
  sheet.getRange(`D4`).setValue('=CHAR(HEX2DEC("1F4AE"))');
  sheet.getRange(`D5`).setValue("interviewer");
};

const onOpen = () => {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("スタンプラリー表設定", [
    { name: "表を更新", functionName: "init" },
  ]);
};
