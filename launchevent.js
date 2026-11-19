/*
 * 送信先確認アドイン：送信ボタン押下時に To/Cc/Bcc の一覧を表示する
 * クラシック版Windowsでも動くよう ES2016 の範囲で記述（async/await・import は使わない）
 * v1.1: どこでエラーが起きても必ず event.completed を呼ぶように修正
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
    var pending = fields.length;

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
            results[field] = result.status === Office.AsyncResultStatus.Succeeded ? result.value : [];
          } catch (e) {
            results[field] = [];
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
        var name = r.displayName && r.displayName !== r.emailAddress
          ? r.displayName + " <" + r.emailAddress + ">"
          : r.emailAddress;
        lines.push(labels[field] + ": " + (ext ? "【社外】" : "") + name);
        total++;
      });
    });

    // 宛先が無い場合はOutlook本体のチェックに任せる
    if (total === 0) {
      complete({ allowEvent: true });
      return;
    }

    var header = "以下の宛先に送信します（" + total + "件"
      + (externalCount > 0 ? "、うち社外 " + externalCount + "件" : "")
      + "）。よろしければ「そのまま送信する」を押してください。\n";

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

    // allowEvent:false + SendMode=PromptUser で、確認ダイアログを毎回表示する
    complete({
      allowEvent: false,
      errorMessage: header + body
    });
  }
}

// マニフェストの FunctionName と関数を対応付ける
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);