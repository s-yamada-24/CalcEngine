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

## 使い方

`CalcEngine` クラスを使用するのが最も簡単です。

```csharp
using CalcEngine;

// 1. インスタンスの作成
var engine = new CalcEngine();

// 2. 値の設定 (仮想テーブル)
engine.SetValue("A1", 10);
engine.SetValue("A2", 20);
engine.SetValue("B1", "Hello");

// 3. 数式の評価
var result1 = engine.Evaluate("=A1 + A2 * 2"); // 50
var result2 = engine.Evaluate("=SUM(A1:A2)");  // 30
var result3 = engine.Evaluate("=IF(A1 > 5, \"OK\", \"NG\")"); // "OK"

// 4. カスタム関数の登録
engine.RegisterFunction("TRIPLE", args => 
{
    return Convert.ToDouble(args[0]) * 3;
});
var result4 = engine.Evaluate("=TRIPLE(A1)"); // 30
```

## サポートしている機能

### 演算子
*   算術: `+`, `-`, `*`, `/`, `^` (べき乗)
*   比較: `=`, `<>`, `<`, `>`, `<=`, `>=`
*   文字列: `&` (連結)

### 標準関数

#### 数学
*   `SUM(range)`
*   `AVERAGE(range)`
*   `MIN(range)`
*   `MAX(range)`
*   `COUNT(range)`
*   `COUNTIF(range, criteria)`
*   `SUMIF(range, criteria, [sum_range])`
*   `ROUND(number, digits)`
*   `ABS(number)`
*   `SQRT(number)`

#### 論理
*   `IF(condition, true_val, false_val)`
*   `AND(logical1, ...)`
*   `OR(logical1, ...)`
*   `NOT(logical)`

#### 文字列
*   `CONCATENATE(text1, ...)`
*   `LEFT(text, num_chars)`
*   `RIGHT(text, num_chars)`
*   `MID(text, start, num_chars)`
*   `LEN(text)`
*   `UPPER(text)`
*   `LOWER(text)`

## エラー処理

以下のエラーコードに対応しており、例外をスローせずに結果としてエラーオブジェクト (`CalcError`) を返します。

*   `#DIV/0!`: ゼロ除算
*   `#REF!`: 無効なセル参照 (存在しないセルへのアクセス)
*   `#VALUE!`: 型変換エラー、引数エラー
*   `#NAME?`: 未定義の関数名
*   `#NUM!`: 数値エラー (負の数の平方根など)

## ライセンス

MIT License
