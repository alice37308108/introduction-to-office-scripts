# 3. Officeスクリプトの基本を学ぼう
***

## 3.1. 変数・定数

### 3.1.1. 変数の宣⾔

数値型の変数numに100を代⼊します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    let num: number;
    num = 100;
}
```

数値型の変数numの宣⾔時に100を代⼊します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    let num: number = 100;
}
```

変数は上書き代⼊ができます。

```tsx
function main(workbook: ExcelScript.Workbook) {
    let num: number;
    num = 100;
    num = 200;
}
```

### 3.1.2. 定数の宣⾔

⽂字列型の変数msgにHelloを代⼊します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const msg: string = "Hello";
}
```

## 3.2. データ型

### 3.2.1. 数値型

数値を扱うときは 数値型（number） を使います。数値型は整数や⼩数を扱うことができます。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const num: number = 100;
    console.log(num); // 100
}
```

数値型は加算（+）減算（-）乗算（*）除算（/）剰余（%）累乗（**）といった演算を⾏うことができます。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 7;
    const y: number = 2;
    console.log(x + y); //9
    console.log(x - y); //5
    console.log(x * y); //14
    console.log(x / y); //3.5
    console.log(x % y); //1
    console.log(x ** y); //49
}
```

### 3.2.2. ⽂字列型

⽂字列を扱うときは ⽂字列型（string） を使います。⽂字列型は⽂字列を扱うことができます。

```tsx
function main(workbook: ExcelScript.Workbook){
 const msg: string = "Hello";
 console.log(msg); // Hello
}
```

⽂字列はプラス記号（+）を使うことで連結できます。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const msg: string = "Hello";
    console.log(msg); // Hello
}
```

テンプレート⽂字列を使ってログを出⼒します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const vegetable: string = "さつまいも";
    const count: number = 3;
    console.log(`${vegetable}を${count}個ください`); //さつまいもを3個ください
}
```

### 3.2.3. 真偽型

真偽型を扱うときは 真偽型（boolean） を使います。真偽型はtrue（真）とfalse（偽）のどちらかを扱うことができます。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const flag: boolean = true;
    console.log(flag); // true
}
```

### 3.2.4. データ型の確認

「100」、「Hello」、「true」のデータ型を確認します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    console.log(typeof (100)); //number
    console.log(typeof ("Hello")); //string
    console.log(typeof (true)); //boolean
}
```

## 3.3. データ型を扱うときの注意点

### 3.3.1. データ型を省略したとき

下の例では変数totalのデータ型を宣⾔していません。

（エラーになります）

```tsx
function main(workbook: ExcelScript.Workbook) {
    const price: number = 100;
    const count: number = 3;
    let total
    total = price * total;
    console.log(total);
}
```

変数totalを数値型で宣⾔します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const price: number = 100;
    const count: number = 3;
    let total: number;
    total = price * count;
    console.log(total); //300
}
```

変数totalに0を代⼊して宣⾔します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const price: number = 100;
    const count: number = 3;
    let total = 0;
    total = price * count;
    console.log(typeof (total)); //number
    console.log(total); //300
}
```

### 3.3.2. 異なるデータ型の値を⼊⼒したとき

数値型として宣⾔した変数numに「Hello」という⽂字列を代⼊しています。

（エラーになります）

```tsx
function main(workbook: ExcelScript.Workbook) {
    let val: number;
    val = 100;
    console.log(val);
    val = "Hello";
    console.log(val);
}
```

## 3.4. 配列・オブジェクト

### 3.4.1. 配列

インデックス0の要素「紅はるか」を取り出します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto: string[] = ["紅はるか", "安納芋", "鳴⾨⾦時"];
    console.log(sweetpoteto[0]); //紅はるか
}
```

インデックス3に「紅あずま」を格納します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto: string[] = ["紅はるか", "安納芋", "鳴⾨⾦時"];
    sweetpoteto[3] = "紅あずま";
    console.log(sweetpoteto); //["紅はるか", "安納芋", "鳴⾨⾦時", "紅あずま"]
}
```

すでにインデックスが存在しているときは上書きとなります。インデックス2の「鳴⾨⾦時」を「五郎島⾦時」に上書きします。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto: string[] = ["紅はるか", "安納芋", "鳴⾨⾦時"];
    sweetpoteto[2] = "五郎島⾦時";
    console.log(sweetpoteto); //["紅はるか", "安納芋", "五郎島⾦時"]
}
```

配列の最後に「シルクスイート」を追加します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto: string[] = ["紅はるか", "安納芋", "鳴⾨⾦時"];
    sweetpoteto.push("シルクスイート");
    console.log(sweetpoteto); //["紅はるか", "安納芋", "鳴⾨⾦時", "シルクスイート"]
}
```

### 3.4.2. オブジェクト

sweetpotetoというオブジェクトに値を格納します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto = {
        カロリー: 140,
        味: "おいしい"
    };
}
```

カロリーと味の値を取り出します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const sweetpoteto = {
        カロリー: 140,
        味: "おいしい"
    };
    sweetpoteto["食物繊維"] = "たくさん"
    console.log(sweetpoteto["食物繊維"]); // たくさん
}
```