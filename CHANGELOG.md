# おまんぼさんイラストゲーム 開発履歴

> **⚠️ このファイルは上書き禁止。追記のみ行うこと。**
> 新しい変更は「## [日付] タイトル」の形式で末尾に追記する。

---

## システム構成

| 要素 | 詳細 |
|------|------|
| フロントエンド | GitHub Pages (`https://omanbosan.github.io/omanbo-illust-game/`) |
| バックエンド | Google Apps Script Web App |
| DB | Google Sheets（`scores` シート・`drawings` シート・`votes` シート） |
| 画像保存 | Google Drive（フォルダID: `1gXwYsxpKBqZ6tMKdNyOPGVvtQTcV1QgN`） |
| GAS URL | `https://script.google.com/macros/s/AKfycby0q8MVbN-Dk_fio8XOAqZW4XnxJC4_RNw4__XQOt57RgQUfAEjqkirPBHCHlZ30D4JoA/exec` |
| GASプロジェクト | `https://script.google.com/home/projects/1TeBoLRDrBZe9QthYYyjTeBpfdyGjFVG2ESgImPcqogTEnu1eKW0iyAlk/edit` |
| スプレッドシート | ID: `1nkGoMoNKc4opnD-Z_Z41I_4re8Brz7lFyBdNXJfpyIw` |

---

## ファイル構成

```
game.html       - メインゲーム画面
admin.html      - 管理者承認画面
gas-backend.gs  - Google Apps Script バックエンド
index.html      - トップページ
```

---

## 採点ロジック（最新）

| モード | 計算式 | 満点ライン（IoU） |
|--------|--------|-----------------|
| 鬼モード | `round(20 × min(1, IoU × 3.8))` | IoU ≈ 26% |
| 通常モード | `round(20 × min(1, IoU × 5.46))` | IoU ≈ 18% |
| イージーモード | `round(20 × min(1, recall × 7.0))` | recall ≈ 14% |

- 5パーツ × 20点 + ボーナス40点 = 最高140点
- 鬼モードのみ採点図鑑への登録・参加番号が発行される

---

## 変更履歴

---

### [2025年頃〜] 初期機能（compaction前の記録）

- ゲーム基本機能（お絵描きチャレンジ・採点チャレンジ）
- イージー / 通常 / 鬼モード の3段階
- Google Drive 画像保存
- 月次ランキング・アーカイブ機能
- 日次まとめメール通知

---

### [セッション1] iPad対応・UI改善

**コミット:** `f7567b8` `fdfcc53`

- **iPad 描画線ズレ修正**: `fh-canvas` の座標計算を `play-canvas.getBoundingClientRect() + _gameScale` 方式に変更
- **iPad ボタン効かない修正**: `_fhCanvasInitialized` フラグでイベントリスナー重複登録を防止
- **iPad キャンバス拡大**: 画面幅600px以上の場合、最大700pxまでキャンバスを拡大表示

```javascript
function applyGameScale() {
  const isTablet = window.innerWidth >= 600;
  const maxAllowed = isTablet ? Math.min(700, window.innerHeight - 220) : 460;
  _gameScale = Math.min(maxAllowed, window.innerWidth - 24) / 460;
}
```

---

### [セッション2] 著作権・パロディ表示変更

**コミット:** `dc4cbb6` `d710922` `8c89ea7` `350414b`

- 「人気のキャラクターは商品化検討中」記載を削除
- パロディ系キャラへの注意書きを追加
- 表現変遷: `「アンパンマン・ドラえもん」` → `「テレビや本に登場するキャラクター」` → **`「アニメキャラクターなどに似せた絵」`**（最終）
- 図鑑に「管理者の思い出図鑑に保存」メッセージ追加
- お友達図鑑に著作権注意書き追加

---

### [セッション3] 管理画面 ログイン保持

**コミット:** `7c420d6`

- localStorage でログイン状態を**30日間**保持
- 「ログアウト」ボタン追加
- セッションキー: `omanbo_admin_session`

---

### [セッション4] 採点チャレンジ 登録フォーム制御

**コミット:** `c5f5af3` `24f959f`

- イージー / 通常モードは登録フォームを**非表示**（鬼モードのみ登録可能）
- 鬼モード登録: Instagram・フォロー確認・商業利用同意の**3つすべて必須**
- 同意しないと登録ボタンが押せない

---

### [セッション5] 鬼モード 輪郭チラ見せ機能

**コミット:** `a492e45` `544a45c`

- 鬼モード開始時、**1パーツ目のみ 0.5秒だけ**輪郭位置を表示
- 2パーツ目以降は表示なし（簡単になりすぎるため）

```javascript
if (index === 0) {
  await drawAnswerImageGuide(ctx);
  await new Promise(res => setTimeout(res, 500));
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}
```

---

### [セッション6] 採点図鑑 ロード改善

**コミット:** `24c17d3`

- ページ読み込み2.5秒後にバックグラウンドで先読み（prefetch）
- スケルトンローディング表示（6枚のシマーカード）
- 20件ずつページネーション（「もっと見る」ボタン）

---

### [セッション7] 参加番号システム導入

**コミット:** `009a783` `babe48f` `a4ec09f`

- 投稿時に**ランダム5桁の参加番号（10000〜99999）**を自動生成
- 生成した画像の右下に番号をスタンプして保存
- GASの `scores` / `drawings` シートに記録
- Instagram は任意入力に変更（番号で当選管理）
- 不正防止: 番号はランダムのため推測困難

```javascript
function generateEntryId() {
  return String(Math.floor(10000 + Math.random() * 90000));
}
```

---

### [セッション8] 鬼モード 採点multiplier調整

**コミット:** `2cc3c59`

- 採点が厳しすぎるとの判断で調整
- `2.94` → **`3.8`**（満点ラインが IoU 34% → 26% に緩和）

```javascript
const multipliers = { hard: 3.8, normal: 5.46 };
```

---

### [セッション9] 管理画面 参加番号・スプシ情報表示

**コミット:** `da9daea` `6275b6f`

- GAS: `getPendingScores` / `getPending` が `entryId` を返すように修正
- GAS: `getZukan` / `getZukanScoring` の `.reverse()` を削除 → **古い順（No.1が最初）** に変更
- 管理画面カードに参加番号をオレンジの大きなバッジで表示
- スプシの全情報をテーブル形式で表示（名前・Instagram・日付・難易度など）
- **参加番号での検索ボックス**追加

---

### [2026-06-13] 投票数のGAS記録・管理画面投票ランキング・ボタンUI整理

**コミット:** `177a4b9`

**GAS変更（要デプロイ）:**
- `votes` シート追加（列: no, 村人名, 票数, 最終投票日）
- `castVote(data)` 関数追加: 投票時に votes シートに記録（upsert）
- `getVoteRanking()` 関数追加: 票数降順で返す
- `doPost`: `type === 'vote'` ルーティング追加
- `doGet`: `type === 'voteRanking'` ルーティング追加

**game.html変更:**
- `castZukanVote()`: 投票時に GAS へも fire-and-forget で送信

**admin.html変更:**
- 「🏆 投票」モードボタン追加
- 投票ランキングテーブル表示（バーチャート付き・票数降順）
- ボタンスタイルを**透明背景＋枠線**のシンプルなデザインに統一

---

### [2026-06-13] 管理画面 スプシ情報表示強化

**コミット:** `6275b6f`

- 参加番号を大きなオレンジバッジで表示（採点・お友達両カード）
- 全スプシ情報をテーブル形式で表示
  - 採点カード: 名前 / Instagram / 投稿日 / 難易度
  - お友達カード: ペンネーム / Instagram / IG表示設定 / 投稿日
- 参加番号での絞り込み検索ボックス追加

---

### [2026-06-13] 投票取り消し・再投票機能追加

**コミット:** `3a3601c`

**GAS変更（要デプロイ）:**
- `castVote()` に `delta` パラメータ追加（+1 or -1）
- `doPost`: `type === 'unvote'` で `delta: -1` を渡す

**game.html変更:**
- 投票済みボタンを **「❤️ 取り消す」** に変更
- タップすると投票取り消し → 残り票数が戻り別キャラに再投票可能
- GAS にも取り消し（-1票）を送信

---

## GAS デプロイ履歴メモ

| 日付 | 理由 |
|------|------|
| セッション序盤 | 初期機能一式 |
| 2026-06-13午前 | entryId保存・図鑑順序修正・getPending系修正 |
| 2026-06-13午後 | votes シート・castVote・getVoteRanking 追加 |
| 2026-06-13夕方 | unvote（取り消し）対応 |

---

## よく使うリンク

- [スプレッドシート](https://docs.google.com/spreadsheets/d/1nkGoMoNKc4opnD-Z_Z41I_4re8Brz7lFyBdNXJfpyIw/edit)
- [Google Drive フォルダ](https://drive.google.com/drive/folders/1gXwYsxpKBqZ6tMKdNyOPGVvtQTcV1QgN)
- [GAS プロジェクト](https://script.google.com/home/projects/1TeBoLRDrBZe9QthYYyjTeBpfdyGjFVG2ESgImPcqogTEnu1eKW0iyAlk/edit)
- [本番サイト](https://omanbosan.github.io/omanbo-illust-game/)
- [管理画面](https://omanbosan.github.io/omanbo-illust-game/admin.html)

---

### [2026-06-27] 採点基準説明を初期画面に追加

**コミット:** `9199e9c`

- 採点チャレンジ開始前の画面に「📐 採点基準」説明ボックスを追加
- 5パーツ × 20点 = 100点満点の構成を明示
- 点数帯ごとの評価ラベルを表示（🌟18〜20点 / 👍14〜17点 / 😅8〜13点 / 💦0〜7点）

---

### [2026-06-27] 採点システム全面刷新・フリースタイルお絵描き追加

**コミット:** `974695a` `10f2716`

**採点システム刷新:**
- ボーナス40点廃止 → **純粋に5パーツ × 20点 = 100点満点**
- 採点結果にパーツ別テーブル表示（得点・バーチャート・評価絵文字）
- 鬼モード multiplier: `4.0`（満点ラインIoU=25%、上位約5%が95点超）
- 通常モード multiplier: `9.0`

**採点バグ修正 (コミット: `1271eb7`):**
- `scoreFreehandOverlap` の `lineWidth` を `Math.max(20, cfg.h*0.28)` → **`8`** に修正
- ガバ描きでも高得点が出てしまう問題を解消

**お絵描きチャレンジ フリースタイル追加:**
- ⭕丸型 / 🔲フリースタイル の切り替えボタンを追加
- フリースタイル時は正方形キャンバス・枠線のみ（SVG輪郭非表示）
- 保存画像もshapeに応じて丸/四角で書き出し

---

### [2026-06-27] 景品テキスト変更・地獄モード削除

- 鬼モード景品: `レザー調キーホルダー` → `コースター`
- 地獄モード（新キャラ）を試験追加したが品質不足のため完全削除


---

### [2026-06-27] 採点multiplierを管理画面から設定可能に

**コミット:** `67c0806`

**GAS変更（要デプロイ）:**
- `config` シート追加（列: key / value）
- `getConfig()` 関数: `hardMult` / `normalMult` を返す（デフォルト 2.0 / 4.5）
- `setConfig(data)` 関数: config シートに upsert
- `doGet`: `type === 'config'` ルーティング追加
- `doPost`: `type === 'setConfig'` ルーティング追加

**game.html変更:**
- 起動時に GAS から `_scoringMultipliers` を非同期取得
- GAS 取得失敗時はデフォルト値（hard: 2.0, normal: 4.5）を使用
- `scoreFreehandOverlap` が `_scoringMultipliers` を参照するように変更

**admin.html変更:**
- 「⚙️ 採点設定」モードボタン追加
- 鬼モード / 通常モード の multiplier 入力欄
- 入力値から「満点ラインとなるIoU%」をリアルタイム表示
- 💾 保存 / 🎲 ランダム設定（鬼: 1.5〜3.5 / 通常: 3.0〜7.0）/ ↩ デフォルト リセット ボタン
