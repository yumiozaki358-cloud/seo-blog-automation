/**
 * SEOブログ記事自動生成システム - GAS本体
 *
 * スプレッドシート列構成（1行目はヘッダー想定）
 *   A列: キーワード         (メインキーワード)
 *   B列: サブキーワード     (SEO関連語・任意)
 *   C列: ステータス         (未生成 / 生成中 / 生成済み / 投稿済み)
 *   D列: タイトル           (AI生成後のタイトル)
 *   E列: WordPress URL      (下書き投稿後のURL・実投稿しない場合は空のまま)
 *   F列: 本文               (AI生成後の記事本文)
 *   G列: メタディスクリプション (検索結果に表示される説明文・120字以内)
 */

const COL = {
  KEYWORD: 1,
  SUB_KEYWORD: 2,
  STATUS: 3,
  TITLE: 4,
  WP_URL: 5,
  BODY: 6,
  META: 7,
};

const STATUS = {
  NOT_GENERATED: '未生成',
  GENERATING: '生成中',
  GENERATED: '生成済み',
  POSTING: '投稿中',
  POSTED: '投稿済み',
  VERIFIED: '確認済み（未投稿）',
};

// Claude APIの設定。max_tokensは「本文2,000〜3,000字＋メタ+JSON構造分」を
// 十分にカバーしつつ、GASの6分制限内で確実に完了する値として設定
const CLAUDE_CONFIG = {
  MODEL: 'claude-sonnet-5', // 利用可能な最新モデル名はAnthropicドキュメントで要確認
  MAX_TOKENS: 8192, // 文字数下限を引き上げた際、4096では出力が途中で打ち切られJSONが不完全になったため増量
  API_URL: 'https://api.anthropic.com/v1/messages',
};

// クライアントから共有されたサンプル記事（トーン・構成の見本として使用）
// 本来はコードに直書きせず専用シートやプロパティに切り出すのが望ましいが、
// 今回はステップの範囲内としてコード内に定数化している
const SAMPLE_ARTICLE = `# 「勤怠管理はエクセルで十分」は本当か？担当者が見落としがちな3つの落とし穴

「うちはまだ社員数もそこまで多くないし、勤怠管理はエクセルで回せている」——そう考えている労務・総務担当の方は少なくありません。確かにエクセルは手軽で、追加コストもかからない優秀なツールです。ただ、会社の成長や働き方の変化に合わせて、いつのまにか"見えないコスト"が膨らんでいることも事実です。この記事では、エクセル勤怠を続けるうえで担当者が見落としがちな3つの落とし穴を、現場目線で整理します。

## 落とし穴①：集計ミスは「気づけないまま」積み上がる
エクセル勤怠で最も多いトラブルが、関数のズレやコピペミスによる集計誤りです。1人分の残業時間が数分ずれるだけなら気づきにくいですが、それが数十人・十数か月分積み重なると、未払い残業として後から大きな金額で表面化することがあります。

## 落とし穴②：法改正への対応が、どうしても後手に回る
働き方関連の法改正は、ここ数年だけでも「有給休暇の年5日取得義務」「36協定の残業上限規制」など立て続けに行われています。エクセル運用の場合、改正のたびに数式やシートを自前で組み直す必要があり、対応が遅れるほど労務リスクが高まります。

## 落とし穴③：テレワークや多様な働き方に追いつけない
テレワーク・時差出勤・中抜けといった柔軟な働き方が広がると、「いつ・どこで・どれだけ働いたか」を正確に記録すること自体が難しくなります。自己申告ベースのエクセルでは、打刻漏れや申告忘れが増えます。

## まとめ：エクセルが悪いのではなく、「フェーズ」の問題
大切なのは、自社の人数・働き方が「エクセルで回せるフェーズ」を超えていないかを定期的に見直すことです。`;

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('記事生成')
    .addItem('選択行から記事を生成', 'generateArticleForSelectedRow')
    .addItem('選択行をWordPressへ下書き投稿', 'postSelectedRowToWordPress')
    .addSeparator()
    .addItem('Claude APIキーを設定', 'setApiKey_')
    .addItem('WordPress接続情報を設定', 'setWordPressConfig_')
    .addToUi();
}

/**
 * Claude APIキーをスクリプトプロパティに保存する（コードへの直書きを避けるため）。
 */
function setApiKey_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Claude APIキーを入力してください', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = result.getResponseText().trim();
  if (!apiKey) {
    ui.alert('APIキーが空のため保存しませんでした。');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', apiKey);
  ui.alert('APIキーを保存しました。');
}

function getApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) {
    throw new Error('Claude APIキーが未設定です。メニューの「Claude APIキーを設定」から登録してください。');
  }
  return apiKey;
}

/**
 * WordPressのURL・ユーザー名・アプリケーションパスワードをスクリプトプロパティに保存する。
 * アプリケーションパスワードは「ユーザー→プロフィール→アプリケーションパスワード」で発行する。
 */
function setWordPressConfig_() {
  const ui = SpreadsheetApp.getUi();

  const urlResult = ui.prompt('WordPressのURLを入力してください（例: https://example-saas.co.jp）', ui.ButtonSet.OK_CANCEL);
  if (urlResult.getSelectedButton() !== ui.Button.OK) return;
  const url = urlResult.getResponseText().trim().replace(/\/$/, ''); // 末尾のスラッシュを除去

  const userResult = ui.prompt('WordPressのユーザー名を入力してください', ui.ButtonSet.OK_CANCEL);
  if (userResult.getSelectedButton() !== ui.Button.OK) return;
  const user = userResult.getResponseText().trim();

  const passResult = ui.prompt('アプリケーションパスワードを入力してください（スペースが入っていてもそのまま貼り付けてください）', ui.ButtonSet.OK_CANCEL);
  if (passResult.getSelectedButton() !== ui.Button.OK) return;
  const appPassword = passResult.getResponseText().trim();

  if (!url || !user || !appPassword) {
    ui.alert('入力に不足があるため保存しませんでした。');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('WP_URL', url);
  props.setProperty('WP_USER', user);
  props.setProperty('WP_APP_PASSWORD', appPassword);
  ui.alert('WordPress接続情報を保存しました。');
}

/**
 * WordPress接続情報を取得する。未設定の場合は例外を投げず null を返す
 * （WordPress環境がない受講生でも、確認モードで動作を継続できるようにするため）。
 */
function getWordPressConfig_() {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('WP_URL');
  const user = props.getProperty('WP_USER');
  const appPassword = props.getProperty('WP_APP_PASSWORD');

  if (!url || !user || !appPassword) {
    return null;
  }
  return { url: url, user: user, appPassword: appPassword };
}

/**
 * WordPressに送信するペイロードを組み立てる。
 * 実投稿・確認モードの両方から共通で使う。
 */
function buildWordPressPayload_(title, content, excerpt) {
  return {
    title: title,
    content: content,
    status: 'draft', // 絶対に'publish'にしない（提案書で約束した鉄則）
    excerpt: excerpt,
  };
}

/**
 * メニューから呼ばれるエントリーポイント。選択行のタイトル・本文・メタディスクリプションを
 * WordPressに「下書き」として投稿する。
 */
function postSelectedRowToWordPress() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();

  if (row === 1) {
    ui.alert('ヘッダー行ではなく、記事が生成済みの行を選択してください。');
    return;
  }

  const status = sheet.getRange(row, COL.STATUS).getValue();
  const title = sheet.getRange(row, COL.TITLE).getValue();
  const body = sheet.getRange(row, COL.BODY).getValue();
  const meta = sheet.getRange(row, COL.META).getValue();

  if (!title || !body) {
    ui.alert('タイトルまたは本文が空です。先に「選択行から記事を生成」を実行してください。');
    return;
  }

  if (status === STATUS.POSTED) {
    const response = ui.alert(
      'この行は既に「投稿済み」です。再度下書き投稿しますか？（WordPress側に別の下書きが新規作成されます）',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  } else if (status !== STATUS.GENERATED) {
    const response = ui.alert(
      `ステータスが「${status || '未生成'}」です。記事生成が完了していない可能性がありますが、続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  setRowStatus_(sheet, row, STATUS.POSTING);

  try {
    const config = getWordPressConfig_();

    if (config) {
      // WordPress接続情報あり: 実際にAPIへ投稿する
      const result = postToWordPressDraft_(title, body, meta, config);
      sheet.getRange(row, COL.WP_URL).setValue(result.editLink);
      setRowStatus_(sheet, row, STATUS.POSTED);
      Logger.log(`WordPress下書き投稿成功: id=${result.id}, editLink=${result.editLink}`);
      ui.alert('WordPressに下書きとして投稿しました。公開前に管理画面で内容をご確認ください。');
    } else {
      // WordPress接続情報なし: 実際には投稿せず、送信予定のペイロードをログに出すだけの確認モード
      const payload = buildWordPressPayload_(title, body, meta);
      Logger.log('--- WordPress確認モード: 実際には投稿せず、送信予定のペイロードのみ表示します ---');
      Logger.log(JSON.stringify(payload, null, 2));

      sheet.getRange(row, COL.WP_URL).setValue('(未投稿・実行ログでペイロード確認済み)');
      setRowStatus_(sheet, row, STATUS.VERIFIED);
      ui.alert(
        'WordPress接続情報が未設定のため、実際の投稿は行いませんでした。\n' +
        '送信予定のペイロード（title / content / status:"draft" / excerpt）は、\n' +
        'Apps Scriptエディタの「実行数」またはログから確認できます。'
      );
    }
  } catch (error) {
    // 失敗時は「生成済み」に戻す。記事データ自体は無事なので、
    // WordPress投稿だけをやり直せる状態を保つ（再生成させないための設計）
    setRowStatus_(sheet, row, STATUS.GENERATED);
    Logger.log(`WordPress投稿エラー: ${error.message}`);
    ui.alert(`WordPressへの投稿中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * WordPress REST APIに下書き投稿を行う（実際にAPIを呼ぶのはこの関数のみ）。
 * @return {{id: number, editLink: string, previewLink: string}}
 */
function postToWordPressDraft_(title, content, excerpt, config) {
  const endpoint = `${config.url}/wp-json/wp/v2/posts`;
  const payload = buildWordPressPayload_(title, content, excerpt);

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(`${config.user}:${config.appPassword}`),
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const statusCode = response.getResponseCode();
  const rawText = response.getContentText();

  if (statusCode !== 201 && statusCode !== 200) {
    throw new Error(`WordPress REST API呼び出し失敗（HTTP ${statusCode}）: ${rawText}`);
  }

  const json = JSON.parse(rawText);

  // 万が一publishとして投稿されてしまった場合、即座に検知できるようにする安全弁
  if (json.status !== 'draft') {
    throw new Error(`想定外のステータスで投稿されました（status: ${json.status}）。WordPress側の投稿設定をご確認ください。`);
  }

  // 下書きのlinkフィールドはそのままでは閲覧できないことが多いため、
  // 管理画面の編集URLを組み立てて記録する
  const editLink = `${config.url}/wp-admin/post.php?post=${json.id}&action=edit`;

  return { id: json.id, editLink: editLink, previewLink: json.link };
}

/**
 * メニューから呼ばれるエントリーポイント。
 */
function generateArticleForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();

  if (row === 1) {
    ui.alert('ヘッダー行ではなく、キーワードが入力されている行を選択してください。');
    return;
  }

  const keyword = sheet.getRange(row, COL.KEYWORD).getValue();
  const subKeyword = sheet.getRange(row, COL.SUB_KEYWORD).getValue();
  const currentStatus = sheet.getRange(row, COL.STATUS).getValue();

  if (!keyword) {
    ui.alert('キーワードが空です。A列にキーワードを入力してから実行してください。');
    return;
  }

  if (currentStatus === STATUS.GENERATED || currentStatus === STATUS.POSTED) {
    const response = ui.alert(
      `このキーワードは既に「${currentStatus}」です。再生成しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  setRowStatus_(sheet, row, STATUS.GENERATING);

  try {
    const draft = generateArticleDraft_(keyword, subKeyword);

    sheet.getRange(row, COL.TITLE).setValue(draft.title);
    sheet.getRange(row, COL.BODY).setValue(draft.body);
    sheet.getRange(row, COL.META).setValue(draft.meta_description);
    setRowStatus_(sheet, row, STATUS.GENERATED);

    Logger.log(`記事生成成功: keyword=${keyword}, title=${draft.title}`);
  } catch (error) {
    setRowStatus_(sheet, row, STATUS.NOT_GENERATED);
    Logger.log(`エラー: ${error.message}`);
    ui.alert(`記事生成中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * Claude APIを呼び出し、記事下書き（タイトル・本文・メタディスクリプション）を生成する。
 * @return {{title: string, body: string, meta_description: string}}
 */
function generateArticleDraft_(keyword, subKeyword) {
  const prompt = buildPrompt_(keyword, subKeyword);
  const apiResult = callClaudeApi_(prompt);

  if (apiResult.stopReason === 'max_tokens') {
    throw new Error(
      '生成が途中で打ち切られました（max_tokens到達）。CLAUDE_CONFIG.MAX_TOKENSをさらに増やすか、' +
      '文字数の目標をやや控えめにすることを検討してください。'
    );
  }

  return parseArticleJson_(apiResult.text);
}

/**
 * プロンプト設計の4要素（役割定義／サンプル記事／構成テンプレート／制約条件）を
 * 組み立てる。JSON形式での出力を明示的に指示する。
 */
function buildPrompt_(keyword, subKeyword) {
  return `あなたはBtoB SaaSのマーケティングライターです。中小企業の経営者・IT担当者に向けて、SEOを意識したブログ記事の下書きを作成してください。

# 参考記事（トーン・文体の見本。表現はそのまま使わず、あくまでトーン参考として）
${SAMPLE_ARTICLE}

# 今回書く記事のキーワード
メインキーワード: ${keyword}
サブキーワード: ${subKeyword || 'なし'}

# 記事構成テンプレート
タイトル → 導入（400字程度） → H2見出し×3（各600〜700字程度） → まとめ（300字程度）

# 制約条件（必須）
- 本文全体の文字数は必ず2,400字以上、3,000字以内にすること。2,000字を下回ることは絶対に避け、内容が薄くなりそうな場合は各見出しの説明・具体例・背景説明を追加して厚みを持たせること
- 「程度」はあくまで目安であり、下限（2,400字）を満たすことを最優先すること
- SEOキーワードを不自然にならない範囲で本文に含める
- 参考記事の構成・トーンは参照してよいが、文章表現はゼロから独自に書くこと（他サイトの文章をそのまま流用しない）
- 事実確認が必要な数値・法令名などは断定しすぎず、一般的な傾向として書くこと

# 出力形式
以下のJSON形式のみを出力してください。前置き、説明文、Markdownのコードフェンス（\`\`\`）は一切不要です。
{
  "title": "記事タイトル",
  "body": "記事本文（Markdown形式の見出しを含む）",
  "meta_description": "検索結果に表示される説明文（120字以内）"
}`;
}

/**
 * Claude APIにリクエストを送信し、生成テキストを返す。
 */
function callClaudeApi_(prompt) {
  const payload = {
    model: CLAUDE_CONFIG.MODEL,
    max_tokens: CLAUDE_CONFIG.MAX_TOKENS,
    messages: [{ role: 'user', content: prompt }],
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': getApiKey_(),
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(CLAUDE_CONFIG.API_URL, options);
  const statusCode = response.getResponseCode();
  const rawText = response.getContentText();

  if (statusCode !== 200) {
    throw new Error(`Claude API呼び出し失敗（HTTP ${statusCode}）: ${rawText}`);
  }

  const json = JSON.parse(rawText);

  if (!json.content || !json.content.length) {
    throw new Error(`Claude APIレスポンスにcontentがありません: ${rawText}`);
  }

  // content[0]を決め打ちせず、typeが"text"のブロックを探す
  // （拡張思考モードなどでは先頭ブロックがtext以外になることがあるため）
  const textBlock = json.content.find(function (block) {
    return block.type === 'text';
  });

  if (!textBlock || !textBlock.text) {
    throw new Error(`Claude APIレスポンスにtextブロックが見つかりません: ${rawText}`);
  }

  // max_tokens到達で本文が途中で切れていないかを、呼び出し元で判定できるようにする
  if (json.stop_reason === 'max_tokens') {
    Logger.log('警告: stop_reason=max_tokens - 出力が途中で打ち切られている可能性があります');
  }

  return { text: textBlock.text, stopReason: json.stop_reason };
}

/**
 * Claudeからの応答テキストをJSONとしてパースする。
 * 万一コードフェンスが混入していた場合に備えて簡易的に除去する。
 */
function parseArticleJson_(responseText) {
  const cleaned = responseText.replace(/```json|```/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error(`Claudeの応答をJSONとして解釈できませんでした: ${cleaned.substring(0, 200)}...`);
  }
}

function setRowStatus_(sheet, row, status) {
  sheet.getRange(row, COL.STATUS).setValue(status);
  SpreadsheetApp.flush();
}
