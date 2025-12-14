# CalcEngine

C# (.NET 8) で実装された、Excelライクな数式評価ライブラリです。
仮想テーブル上の値を参照して、文字列で渡された計算式をパース・評価します。

## 特徴

*   **Excel互換の数式構文**: 四則演算、比較演算、文字列連結、関数呼び出し、ネストした括弧に対応。
*   **仮想テーブル**: 値のみを保持する軽量なテーブル構造。セルアドレス (`A1`, `B2` 等) で値を管理。
*   **Excel互換のエラー処理**: `#REF!`, `#DIV/0!`, `#VALUE!` などのエラーコードをサポートし、計算過程で伝播します。
*   **高速化**: パース済みの構文木 (AST) をキャッシュし、同じ数式の再評価を高速化。
*   **拡張性**: カスタム関数を簡単に登録可能。

## インストール

このプロジェクトはクラスライブラリとして提供されています。ソリューションに追加して参照してください。

```bash
# プロジェクトに参照を追加
dotnet add reference path/to/CalcEngine/CalcEngine.csproj
```

## 基本的な使い方

### 1. 簡単な計算

`CalcEngine` クラスを使用するのが最も簡単です。

```csharp
using CalcEngine;

// インスタンスの作成
var engine = new CalcEngine();

// 値の設定 (仮想テーブル)
engine.SetValue("A1", 10);
engine.SetValue("A2", 20);

// 数式の評価
var result = engine.Evaluate("=A1 + A2"); // 30.0
Console.WriteLine(result);
```

### 2. 様々な演算

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 10);
engine.SetValue("A2", 20);
engine.SetValue("B1", "Hello");
engine.SetValue("B2", "World");

// 算術演算
var result1 = engine.Evaluate("=A1 + A2 * 2");     // 50.0
var result2 = engine.Evaluate("=(A1 + A2) * 2");   // 60.0
var result3 = engine.Evaluate("=A2 / A1");         // 2.0
var result4 = engine.Evaluate("=2 ^ 3");           // 8.0 (べき乗)

// 比較演算
var result5 = engine.Evaluate("=A1 > 5");          // true
var result6 = engine.Evaluate("=A1 = A2");         // false
var result7 = engine.Evaluate("=A1 <> A2");        // true

// 文字列連結
var result8 = engine.Evaluate("=B1 & \" \" & B2"); // "Hello World"
```

### 3. セル範囲の参照

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 10);
engine.SetValue("A2", 20);
engine.SetValue("A3", 30);

// 範囲を使った関数
var sum = engine.Evaluate("=SUM(A1:A3)");         // 60.0
var avg = engine.Evaluate("=AVERAGE(A1:A3)");     // 20.0
var max = engine.Evaluate("=MAX(A1:A3)");         // 30.0
var min = engine.Evaluate("=MIN(A1:A3)");         // 10.0
var count = engine.Evaluate("=COUNT(A1:A3)");     // 3.0
```

### 4. 条件付き関数

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 10);
engine.SetValue("A2", 20);
engine.SetValue("A3", 10);

// IF関数
var result1 = engine.Evaluate("=IF(A1 > 5, \"大きい\", \"小さい\")"); // "大きい"

// 論理関数
var result2 = engine.Evaluate("=AND(A1 > 5, A1 < 20)");  // true
var result3 = engine.Evaluate("=OR(A1 > 100, A2 > 15)"); // true
var result4 = engine.Evaluate("=NOT(A1 > 100)");         // true

// COUNTIF / SUMIF
var result5 = engine.Evaluate("=COUNTIF(A1:A3, 10)");    // 2.0
var result6 = engine.Evaluate("=SUMIF(A1:A3, 10)");      // 20.0 (10 + 10)
var result7 = engine.Evaluate("=COUNTIF(A1:A3, \">15\")"); // 1.0
```

### 5. 文字列関数

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", "Hello");
engine.SetValue("A2", "World");
engine.SetValue("A3", "  Trim Me  ");

// 文字列操作
var result1 = engine.Evaluate("=CONCATENATE(A1, \" \", A2)"); // "Hello World"
var result2 = engine.Evaluate("=LEFT(A1, 3)");                // "Hel"
var result3 = engine.Evaluate("=RIGHT(A1, 2)");               // "lo"
var result4 = engine.Evaluate("=MID(A1, 2, 3)");              // "ell" (1-indexed)
var result5 = engine.Evaluate("=LEN(A1)");                    // 5.0
var result6 = engine.Evaluate("=UPPER(A1)");                  // "HELLO"
var result7 = engine.Evaluate("=LOWER(A1)");                  // "hello"
var result8 = engine.Evaluate("=TRIM(A3)");                   // "Trim Me"
```

### 6. 数学関数

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 3.14159);
engine.SetValue("A2", -5);
engine.SetValue("A3", 16);

// 数学関数
var result1 = engine.Evaluate("=ROUND(A1, 2)");    // 3.14
var result2 = engine.Evaluate("=ABS(A2)");         // 5.0
var result3 = engine.Evaluate("=SQRT(A3)");        // 4.0
var result4 = engine.Evaluate("=MOD(10, 3)");      // 1.0
```

## 高度な使い方

### カスタム関数の登録

独自の関数を定義して使用できます。

```csharp
var engine = new CalcEngine();

// 単純なカスタム関数
engine.RegisterFunction("TRIPLE", args => 
{
    return Convert.ToDouble(args[0]) * 3;
});

engine.SetValue("A1", 5);
var result1 = engine.Evaluate("=TRIPLE(A1)"); // 15.0

// 複数引数を取る関数
engine.RegisterFunction("MULTIPLY", args => 
{
    double result = 1;
    foreach (var arg in args)
    {
        result *= Convert.ToDouble(arg);
    }
    return result;
});

var result2 = engine.Evaluate("=MULTIPLY(2, 3, 4)"); // 24.0

// エラーを返す関数
engine.RegisterFunction("SAFE_DIVIDE", args => 
{
    var numerator = Convert.ToDouble(args[0]);
    var denominator = Convert.ToDouble(args[1]);
    
    if (denominator == 0)
    {
        return CalcError.Div0;
    }
    return numerator / denominator;
});

var result3 = engine.Evaluate("=SAFE_DIVIDE(10, 2)"); // 5.0
var result4 = engine.Evaluate("=SAFE_DIVIDE(10, 0)"); // CalcError (#DIV/0!)
```

### 値の取得と管理

```csharp
var engine = new CalcEngine();

// 値の設定
engine.SetValue("A1", 100);
engine.SetValue("A2", "テスト");

// 値の取得
var value1 = engine.GetValue("A1"); // 100
var value2 = engine.GetValue("A2"); // "テスト"

// 存在しないセルの取得
var value3 = engine.GetValue("Z99"); // null

// テーブルのクリア
engine.Clear();
```

### 低レベルAPI（VirtualTableとFormulaEvaluatorを直接使用）

より細かい制御が必要な場合は、低レベルAPIを使用できます。

```csharp
using CalcEngine;

// VirtualTableとFormulaEvaluatorを直接使用
var table = new VirtualTable();
table.SetValue("A1", 10);
table.SetValue("A2", 20);

var evaluator = new FormulaEvaluator(table);
var result = evaluator.Evaluate("=A1 + A2"); // 30.0

// カスタム関数の登録
evaluator.FunctionRegistry.Register("DOUBLE", args => 
{
    return Convert.ToDouble(args[0]) * 2;
});

var result2 = evaluator.Evaluate("=DOUBLE(A1)"); // 20.0
```

## エラー処理

CalcEngineは例外をスローせず、エラーを `CalcError` オブジェクトとして返します。

### エラーの種類

*   `#DIV/0!`: ゼロ除算
*   `#REF!`: 無効なセル参照 (存在しないセルへのアクセス)
*   `#VALUE!`: 型変換エラー、引数エラー
*   `#NAME?`: 未定義の関数名
*   `#NUM!`: 数値エラー (負の数の平方根など)
*   `#N/A`: データが利用できない

### エラーハンドリングの例

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 10);

// ゼロ除算
var result1 = engine.Evaluate("=A1 / 0");
if (result1 is CalcError error1)
{
    Console.WriteLine($"エラー: {error1.Code}"); // エラー: #DIV/0!
}

// 存在しないセル参照
var result2 = engine.Evaluate("=B1 + 10");
if (result2 is CalcError error2)
{
    Console.WriteLine($"エラー: {error2.Code}"); // エラー: #REF!
}

// 未定義の関数
var result3 = engine.Evaluate("=UNKNOWN_FUNC(A1)");
if (result3 is CalcError error3)
{
    Console.WriteLine($"エラー: {error3.Code}"); // エラー: #NAME?
}

// エラーの伝播
engine.SetValue("A2", 0);
var result4 = engine.Evaluate("=A1 / A2 + 100");
if (result4 is CalcError error4)
{
    Console.WriteLine($"エラー: {error4.Code}"); // エラー: #DIV/0!
    // エラーは計算全体に伝播します
}
```

## 実用例

### 例1: 売上計算

```csharp
var engine = new CalcEngine();

// 商品データ
engine.SetValue("A1", 1000);  // 単価
engine.SetValue("B1", 5);     // 数量
engine.SetValue("C1", 0.1);   // 消費税率

// 計算
var subtotal = engine.Evaluate("=A1 * B1");           // 5000 (小計)
var tax = engine.Evaluate("=A1 * B1 * C1");           // 500 (税額)
var total = engine.Evaluate("=A1 * B1 * (1 + C1)");   // 5500 (合計)

Console.WriteLine($"小計: {subtotal}円");
Console.WriteLine($"税額: {tax}円");
Console.WriteLine($"合計: {total}円");
```

### 例2: 成績評価

```csharp
var engine = new CalcEngine();

// 学生の点数
engine.SetValue("A1", 85);  // 国語
engine.SetValue("A2", 92);  // 数学
engine.SetValue("A3", 78);  // 英語

// 評価基準を関数で登録
engine.RegisterFunction("GRADE", args => 
{
    var score = Convert.ToDouble(args[0]);
    if (score >= 90) return "A";
    if (score >= 80) return "B";
    if (score >= 70) return "C";
    if (score >= 60) return "D";
    return "F";
});

// 計算
var average = engine.Evaluate("=AVERAGE(A1:A3)");     // 85.0
var grade = engine.Evaluate("=GRADE(AVERAGE(A1:A3))"); // "B"

Console.WriteLine($"平均点: {average}");
Console.WriteLine($"評価: {grade}");
```

### 例3: 在庫管理

```csharp
var engine = new CalcEngine();

// 在庫データ
engine.SetValue("A1", 100);  // 商品A在庫
engine.SetValue("A2", 5);    // 商品B在庫
engine.SetValue("A3", 50);   // 商品C在庫
engine.SetValue("B1", 10);   // 商品A安全在庫
engine.SetValue("B2", 10);   // 商品B安全在庫
engine.SetValue("B3", 10);   // 商品C安全在庫

// 在庫不足チェック
var checkA = engine.Evaluate("=IF(A1 < B1, \"要発注\", \"OK\")"); // "OK"
var checkB = engine.Evaluate("=IF(A2 < B2, \"要発注\", \"OK\")"); // "要発注"
var checkC = engine.Evaluate("=IF(A3 < B3, \"要発注\", \"OK\")"); // "OK"

Console.WriteLine($"商品A: {checkA}");
Console.WriteLine($"商品B: {checkB}");
Console.WriteLine($"商品C: {checkC}");
```

## パフォーマンス

CalcEngineは以下の最適化を実装しています：

*   **ASTキャッシュ**: 同じ数式を複数回評価する場合、パース済みのASTを再利用します。
*   **例外の最小化**: エラー処理に例外を使用せず、`CalcError` オブジェクトを返すことで高速化。

```csharp
var engine = new CalcEngine();
engine.SetValue("A1", 10);

// 初回はパースが実行される
var result1 = engine.Evaluate("=A1 * 2 + 5");

// 2回目以降はキャッシュされたASTを使用（高速）
var result2 = engine.Evaluate("=A1 * 2 + 5");
var result3 = engine.Evaluate("=A1 * 2 + 5");
```

## サポートしている機能

### 演算子
*   算術: `+`, `-`, `*`, `/`, `^` (べき乗)
*   比較: `=`, `<>`, `<`, `>`, `<=`, `>=`
*   文字列: `&` (連結)

### 標準関数

#### 数学
*   `SUM(range)` - 合計
*   `AVERAGE(range)` - 平均
*   `MIN(range)` - 最小値
*   `MAX(range)` - 最大値
*   `COUNT(range)` - 個数
*   `COUNTIF(range, criteria)` - 条件付き個数
*   `SUMIF(range, criteria, [sum_range])` - 条件付き合計
*   `ROUND(number, digits)` - 四捨五入
*   `ABS(number)` - 絶対値
*   `SQRT(number)` - 平方根
*   `MOD(number, divisor)` - 剰余

#### 論理
*   `IF(condition, true_val, false_val)` - 条件分岐
*   `AND(logical1, ...)` - 論理積
*   `OR(logical1, ...)` - 論理和
*   `NOT(logical)` - 論理否定

#### 文字列
*   `CONCATENATE(text1, ...)` - 文字列連結
*   `LEFT(text, num_chars)` - 左から文字列抽出
*   `RIGHT(text, num_chars)` - 右から文字列抽出
*   `MID(text, start, num_chars)` - 中間から文字列抽出
*   `LEN(text)` - 文字列長
*   `UPPER(text)` - 大文字変換
*   `LOWER(text)` - 小文字変換
*   `TRIM(text)` - 空白削除

## テスト

プロジェクトには包括的なユニットテストが含まれています。

```bash
# テストの実行
dotnet test
```

## ライセンス

MIT License
