/*
 * 送信先確認アドイン：送信ボタン押下時に To/Cc/Bcc の一覧を表示する
 * クラシック版Windowsでも動くよう ES2016 の範囲で記述（async/await・import は使わない）
 * v1.1: どこでエラーが起きても必ず event.completed を呼ぶように修正
 * v1.2(調査用): 宛先取得の失敗理由をダイアログとコンソールに出す
 * v1.3(調査用): Office.onReady 後にも associate し、初期化状況と未捕捉エラーを記録
 * v1.4: 表示を短縮（社内=名前<アドレス>、部署名の括弧は省略／社外=アドレスのみ）
 * v1.5: 見出しと宛先一覧の間に空行を入れる
 */

// ダイアログに出せる文字数には上限があるため、余裕を持って切り詰める
var MAX_MESSAGE_LENGTH = 450;

function onMessageSendHandler(event) {
  var done = false;

  // 二重呼び出しを防ぎつつ、必ず1回だけ completed を呼ぶ
  function complete(options) {
    if (done) {
      return;
    }
    done = true;
    try {
      event.completed(options);
    } catch (e) {
      console.error("[送信先確認] completed失敗", e);
      // メッセージが原因で失敗した場合に備え、短い文で再試行
      try {
        event.completed({ allowEvent: false, errorMessage: "宛先を確認してから送信してください。" });
      } catch (e2) {
        console.error("[送信先確認] 再試行も失敗", e2);
      }
    }
  }

  function fail(where, e) {
    console.error("[送信先確認] エラー: " + where, e);
    complete({
      allowEvent: false,
      errorMessage: "宛先一覧を取得できませんでした（" + where + "）。宛先を目で確認してから送信してください。"
    });
  }

  try {
    var item = Office.context.mailbox.item;
    var fields = ["to", "cc", "bcc"];
    var labels = { to: "To", cc: "Cc", bcc: "Bcc" };
    var results = {};
    var errors = [];
    var pending = fields.length;
    console.log("[送信先確認] ハンドラー開始");

    // 自分のドメイン（社外宛ての判定用）。取得できなくても続行する
    var myDomain = "";
    try {
      var myAddress = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
      myDomain = myAddress.indexOf("@") >= 0 ? myAddress.split("@")[1] : "";
    } catch (e) {
      console.warn("[送信先確認] 自分のアドレスを取得できないため社外判定を省略", e);
    }

    fields.forEach(function (field) {
      try {
        item[field].getAsync(function (result) {
          try {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              results[field] = result.value || [];
              console.log("[送信先確認] " + field + " 取得成功: " + results[field].length + "件");
            } else {
              results[field] = [];
              var msg = result.error ? (result.error.code + " " + result.error.message) : "不明";
              errors.push(field + ": " + msg);
              console.error("[送信先確認] " + field + " 取得失敗: " + msg);
            }
          } catch (e) {
            results[field] = [];
            errors.push(field + ": " + e);
          }
          pending--;
          if (pending === 0) {
            try {
              finish();
            } catch (e) {
              fail("一覧作成", e);
            }
          }
        });
      } catch (e) {
        fail(field + "取得", e);
      }
    });
  } catch (e) {
    fail("初期化", e);
  }

  function isExternal(address) {
    var a = (address || "").toLowerCase();
    if (!myDomain || a.indexOf("@") < 0) {
      return false;
    }
    return a.split("@")[1] !== myDomain;
  }

  function finish() {
    var lines = [];
    var total = 0;
    var externalCount = 0;

    fields.forEach(function (field) {
      var list = results[field] || [];
      list.forEach(function (r) {
        var ext = isExternal(r.emailAddress);
        if (ext) {
          externalCount++;
        }
        var name;
        if (ext) {
          // 社外：ドメインの間違いに気づけるよう、アドレスのみ表示
          name = r.emailAddress;
        } else {
          // 社内：表示名から（部署名）などの括弧書きを除き、アドレスを併記
          var dn = (r.displayName || "").replace(/[（(][^（）()]*[）)]/g, "").replace(/[\s　]+$/, "").replace(/^[\s　]+/, "");
          name = dn && dn !== r.emailAddress ? dn + " <" + r.emailAddress + ">" : r.emailAddress;
        }
        lines.push(labels[field] + ": " + (ext ? "【社外】" : "") + name);
        total++;
      });
    });

    // 取得に失敗した欄がある場合は、理由を表示して止める（調査用）
    if (errors.length > 0) {
      complete({
        allowEvent: false,
        errorMessage: "宛先の取得に失敗しました: " + errors.join(" / ")
      });
      return;
    }

    // 宛先が無い場合はOutlook本体のチェックに任せる
    if (total === 0) {
      complete({ allowEvent: true });
      return;
    }

    var header = "宛先を確認してください（" + total + "件"
      + (externalCount > 0 ? "／社外" + externalCount + "件" : "")
      + "）\n\u3000\n"; // 空行が詰められないよう全角スペースだけの行を挟む

    var body = "";
    var shown = 0;
    for (var i = 0; i < lines.length; i++) {
      var next = body + lines[i] + "\n";
      var rest = lines.length - (i + 1);
      var tail = rest > 0 ? "…ほか " + rest + "件" : "";
      if ((header + next + tail).length > MAX_MESSAGE_LENGTH) {
        break;
      }
      body = next;
      shown++;
    }
    if (shown < lines.length) {
      body += "…ほか " + (lines.length - shown) + "件";
    }

    console.log("[送信先確認] ダイアログ表示: " + total + "件");
    // allowEvent:false + SendMode=PromptUser で、確認ダイアログを毎回表示する
    complete({
      allowEvent: false,
      errorMessage: header + body
    });
  }
}

// マニフェストの FunctionName と関数を対応付ける
console.log("[送信先確認] launchevent.js v1.5 読み込み");

// 未捕捉のエラーも記録する
if (typeof window !== "undefined" && window.addEventListener) {
  window.addEventListener("error", function (e) {
    console.error("[送信先確認] 未捕捉エラー: " + e.message + " @" + e.filename + ":" + e.lineno);
  });
  window.addEventListener("unhandledrejection", function (e) {
    console.error("[送信先確認] 未捕捉Promise: " + e.reason);
  });
}

function associateHandler(when) {
  try {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    console.log("[送信先確認] associate 完了 (" + when + ")");
  } catch (e) {
    console.error("[送信先確認] associate 失敗 (" + when + ")", e);
  }
}

// 1回目：読み込み直後（Microsoft公式の書き方）
associateHandler("即時");

// 2回目：Office.js の初期化完了後（初期化が終わっているかの確認も兼ねる）
try {
  Office.onReady(function (info) {
    console.log("[送信先確認] onReady host=" + (info && info.host) + " platform=" + (info && info.platform));
    associateHandler("onReady後");
  });
} catch (e) {
  console.error("[送信先確認] onReady 登録失敗", e);
}