/*
 * 送信先確認アドイン：送信ボタン押下時に To/Cc/Bcc の一覧を表示する
 * クラシック版Windowsでも動くよう ES2016 の範囲で記述（async/await・import は使わない）
 */

// ダイアログに出せる文字数には上限があるため、余裕を持って切り詰める
var MAX_MESSAGE_LENGTH = 450;

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var fields = ["to", "cc", "bcc"];
  var labels = { to: "To", cc: "Cc", bcc: "Bcc" };
  var results = {};
  var pending = fields.length;

  // 自分のドメイン（社外宛ての判定用）
  var myAddress = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  var myDomain = myAddress.indexOf("@") >= 0 ? myAddress.split("@")[1] : "";

  fields.forEach(function (field) {
    item[field].getAsync(function (result) {
      results[field] = result.status === Office.AsyncResultStatus.Succeeded ? result.value : [];
      pending--;
      if (pending === 0) {
        finish();
      }
    });
  });

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
      event.completed({ allowEvent: true });
      return;
    }

    var header = "以下の宛先に送信します（" + total + "件"
      + (externalCount > 0 ? "、うち社外 " + externalCount + "件" : "")
      + "）。よろしければ「このまま送信」を押してください。\n";

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
    event.completed({
      allowEvent: false,
      errorMessage: header + body
    });
  }
}

// マニフェストの FunctionName と関数を対応付ける
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
