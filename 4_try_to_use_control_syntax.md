# 4. Officeスクリプトで制御構文を使ってみよう
***
## 4.1. ⽐較演算⼦・論理演算⼦

### 4.1.1. ⽐較演算⼦

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 5;
    const y: number = 3;
    console.log(x == y); //false
    console.log(x === y); //false
    console.log(x != y); //true
    console.log(x !== y); //true
    console.log(x > y); //true
    console.log(x >= y); //true
    console.log(x < y); //false
    console.log(x <= y); //false
}
```

### 4.1.2. 論理演算⼦

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 5;
    const y: number = 3;
    console.log(x > 3 && y > 3); //false
    console.log(x > 3 || y > 3); //true
    console.log(!(x > 3)); //false
}
```

## 4.2. 条件分岐

### 4.2.1. if⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 80;
    if (x > 70) {
        console.log("合格です！");
    }
}
```

### 4.2.2. if…else⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 50;
    if (x > 70) {
        console.log("合格です！");
    } else {
        console.log("不合格です");
    }
}
```

条件式「x > 70」がtrueの場合は【処理1】を実⾏してif⽂を終了します。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 80;
    if (x > 70) {
        console.log("合格です！");
    } else {
        console.log("不合格です");
    }
}
```

### 4.2.3. if…else if⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 30;
    if (x > 70) {
        console.log("合格です！");
    } else if (x > 50) {
        console.log("半分正解してます！");
    } else {
        console.log("きっといいことがあります。");
    }
}
```

最後のelseブロックは省略可能です。その場合はすべての条件式でfalseとなるため、何も出⼒されません。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const x: number = 30;
    if (x > 70) {
        console.log("合格です！");
    } else if (x > 50) {
        console.log("半分正解してます！");
    }
}
```

### 4.2.4. switch⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const imo: string = "さつまいも";
    switch (imo) {
        case "じゃがいも":
            console.log("じゃがバターにしよう！");
            break;
        case "さつまいも":
            console.log("スイートポテトにしよう！");
            break;
        case "さといも":
            console.log("煮物にしよう！");
            break;
        default:
            console.log("ふかして食べよう！");
    }
}
```

break⽂を省略するとどうなるのでしょうか。

```tsx
function main(workbook: ExcelScript.Workbook) {
    const imo: string = "さつまいも";
    switch (imo) {
        case "じゃがいも":
            console.log("じゃがバターにしよう！");
        case "さつまいも":
            console.log("スイートポテトにしよう！");
        case "さといも":
            console.log("煮物にしよう！");
        default:
            console.log("ふかして食べよう！");
    }
}
```

## 4.3. 繰り返し

### 4.3.1. while⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    let i: number = 1;
    while (i <= 5) {
        console.log(`さつまいもを${i}個食べました`);
        i *= 2;
        45
    }
}
```

### 4.3.2. for⽂

次のスクリプトを実⾏してみましょう。

```tsx
function main(workbook: ExcelScript.Workbook) {
    for (let i = 1; i <= 5; i++) {
        console.log(`さつまいもを${i}個食べました`);
    }
}
```