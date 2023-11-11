# 6. OfficeスクリプトをPowerAutomateで使ってみよう
***
## 6.2. OfficeスクリプトからPower Automateに単⼀の戻り値を渡す

### 6.2.1. 戻り値

戻り値とは関数やメソッドで処理した結果を呼び出し元へ戻す値です。今回は午前のメニューをPower Automateに渡すというスクリプトにしたいので、午前のメニューが戻り値となります。

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook): string {
    const sheet = workbook.getWorksheet("Sheet1");
    const range = sheet.getUsedRange();
    const maxRow: number = range.getRowCount();
    let amMenu: string;
    let pmMenu: string;
    for (let i = 1; i < maxRow; i++) {
        if (sheet.getCell(i, 3).getValue() === "") {
            amMenu = sheet.getCell(i, 1).getValue().toString();
            pmMenu = sheet.getCell(i, 2).getValue().toString();
            console.log(amMenu, pmMenu);
            sheet.getCell(i, 3).setValue("〇");
            break;
        }
    }
    return amMenu;
}
```

## 6.3. OfficeスクリプトからPower Automateに複数の戻り値を渡す

### 6.3.1. interface

午前と午後のメニューを戻り値として渡すスクリプトに変更しましょう。

```tsx
function main(workbook: ExcelScript.Workbook): menu {
    const sheet = workbook.getWorksheet("Sheet1");
    const range = sheet.getUsedRange();
    const maxRow: number = range.getRowCount();
    let amMenu: string;
    let pmMenu: string;
    for (let i = 1; i < maxRow; i++) {
        if (sheet.getCell(i, 3).getValue() === "") {
            amMenu = sheet.getCell(i, 1).getValue().toString();
            pmMenu = sheet.getCell(i, 2).getValue().toString();
            console.log(amMenu, pmMenu);
            sheet.getCell(i, 3).setValue("〇");
            break;
        }
    }
    return {
        午前のメニュー: amMenu,
        午後のメニュー: pmMenu
    };
}
interface menu {
    午前のメニュー: string,
    午後のメニュー: string
}
```

## 6.4. OfficeスクリプトでPower Automateから引数を受け取る

### 6.4.1. 引数

main関数の仮引数を次のように書き換えます。

```tsx
function main(workbook: ExcelScript.Workbook, today: string): menu {
}
```

### 6.4.2. ⽇付操作

これで最後です。少しボリュームがありますが実際に書き換えていきましょう。

```tsx
function main(workbook: ExcelScript.Workbook, today: string): menu {
    const sheet = workbook.getWorksheet("Sheet1");
    const range = sheet.getUsedRange();
    const maxRow: number = range.getRowCount();
    let dt = new Date(today);
    81
    let amMenu: string;
    let pmMenu: string;
    for (let i = 1; i < maxRow; i++) {
        let date = range.getCell(i, 0).getValue() as number;
        let javaScriptDate = new Date(Math.round((date - 25569) * 86400
            * 1000));
        if (dt.getTime() === javaScriptDate.getTime()) {
            amMenu = sheet.getCell(i, 1).getValue().toString();
            pmMenu = sheet.getCell(i, 2).getValue().toString();
            //console.log(amMenu, pmMenu);
            break;
        }
    }
    return {
        午前のメニュー: amMenu,
        午後のメニュー: pmMenu
    };
}
interface menu {
    午前のメニュー: string,
    午後のメニュー: string
}
```
