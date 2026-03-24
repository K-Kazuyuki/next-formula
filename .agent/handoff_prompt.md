# NextFormula - Agent Handoff Prompt

## プロジェクト概要
NextFormulaは、データ（値）とロジック（関数）を完全に分離した新しいアーキテクチャのWebベース・スプレッドシートです。
「AG Grid」でデータを視覚化し、「Monaco Editor」で関数のロジックを管理し、「HyperFormula」を通して双方のTwo-way Sync（双方向同期）を行います。

## 重要なアーキテクチャのルール
- `dataState` は `{ value: any, isFormula: boolean }` のようなメタデータを持つ2次元配列です。
- `formulaState` がすべての計算ロジック情報の単一の情報源（Single Source of Truth）です。プレーンテキストで管理されます。
- `HyperFormula` は関数型ライクに依存関係を自動解決します。行の記述順序は関係なく、循環参照はエラーとして扱います。
- セルの削除時や上書き時は、`formulaState` にある対応する行を完全に削除し、`isFormula` フラグを折ります。
- エラー起因時は、Monaco Editor のマーカーAPIを使って赤い波線を該当行に引き、AG Gridのセルは赤くハイライトします。

## 完了したタスク (Phase 1 & 2)
- Vite + React + TypeScript プロジェクトのセットアップ
- 2ペイン（AG Grid と Monaco Editor）のレイアウト構築
- HyperFormula の導入および `dataState` と `formulaState` からのデータ流し込み・連動計算
- シンタックスエラーや循環参照の検知とハンドリング（Monacoマーカー＋Grid赤色表示）

## これから着手するタスク
1. **Phase 3 (数式バーのUI実装)**: グリッド上部に数式バーを配置し、セルのフォーカス状態と数式テキストを連携させる。
2. **Phase 4 (エイリアス機能)**: `ALIAS [基本攻撃力] = V` や `ALIAS [個別HP] = Z1` といったエイリアス定義をパースし、HyperFormula に渡す前にプレーンなA1参照に置換するプレプロセッサ（パーサー）の実装。

## エージェントへの指示
上記のルールや過去の実装（`Plan.md` および `src/App.tsx`）を読み込み、Phase 3以降のタスクを順次遂行してください。
仕様に疑問点があれば、必ずユーザーに確認してから実装を進めてください。
動作確認をする際は `npm run dev` でローカルサーバーを確認（`http://localhost:5173`）してください。
