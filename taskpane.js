/* 確認画面：宛先・添付を表示し、全項目チェックで確認済みフラグを保存する */
let currentState = null;

function el(tag, props, children) {
  const node = document.createElement(tag);
  Object.assign(node, props || {});
  (children || []).forEach((c) => node.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
  return node;
}

function checkRow(text) {
  const input = el("input", { type: "checkbox" });
  input.addEventListener("change", updateButton);
  return el("label", { className: "check" }, [input, text]);
}

function formatSize(bytes) {
  return bytes >= 1048576 ? (bytes / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(bytes / 1024)) + " KB";
}

function setStatus(text, cls) {
  const s = document.getElementById("status");
  s.textContent = text;
  s.className = cls || "";
}

function updateButton() {
  const boxes = document.querySelectorAll("input[type=checkbox]");
  const allChecked = boxes.length > 0 && Array.from(boxes).every((b) => b.checked);
  document.getElementById("confirm").disabled = !allChecked;
}

// textContent / createTextNode のみ使う（表示名に含まれる文字列をHTMLとして解釈しない）
async function render() {
  setStatus("");
  try {
    currentState = await Okan.collect(Office.context.mailbox.item);
  } catch (e) {
    setStatus("宛先と添付ファイルを読み込めませんでした。画面を開き直してください。", "error");
    return;
  }
  const s = currentState;

  document.getElementById("summary").textContent =
    `宛先 ${s.recipients.length} 件（社外 ${s.externalCount} 件）、添付 ${s.files.length} 件`;

  const rec = document.getElementById("recipients");
  rec.replaceChildren(el("h2", { textContent: "宛先" }));
  if (s.groups.length === 0) rec.appendChild(el("p", { className: "empty", textContent: "宛先がありません。" }));

  s.groups.forEach((g) => {
    const list = el("ul");
    g.recipients.forEach((r) => {
      list.appendChild(
        el("li", {}, [
          el("span", { className: "kind" + (r.kind === "Bcc" ? " bcc" : ""), textContent: r.kind }),
          r.name && r.name !== r.address ? `${r.name} <${r.address}>` : r.address,
        ])
      );
    });
    const head = el("div", { className: "group-head" }, [
      g.domain,
      el("span", { className: "badge", textContent: g.external ? "社外" : "社内" }),
    ]);
    rec.appendChild(
      el("div", { className: "group" + (g.external ? " external" : "") }, [
        head,
        list,
        checkRow(`${g.domain} 宛ての ${g.recipients.length} 件で問題ない`),
      ])
    );
  });

  const att = document.getElementById("attachments");
  att.replaceChildren();
  if (s.files.length > 0) {
    att.appendChild(el("h2", { textContent: "添付ファイル" }));
    const list = el("ul");
    s.files.forEach((f) => list.appendChild(el("li", { textContent: `${f.name}（${formatSize(f.size)}）` })));
    att.appendChild(el("div", { className: "group" }, [list, checkRow("添付ファイルは正しい")]));
  }

  updateButton();
}

async function onConfirm() {
  try {
    // 画面表示後に宛先・添付が変わっていないか再確認してから保存する
    const latest = await Okan.collect(Office.context.mailbox.item);
    if (latest.signature !== currentState.signature) {
      await render();
      setStatus("宛先または添付ファイルが変更されました。もう一度確認してください。", "error");
      return;
    }
    await Okan.setConfirmed(Office.context.mailbox.item, latest.signature);
    setStatus("確認を保存しました。メールの「送信」をもう一度押してください。", "ok");
  } catch (e) {
    setStatus("確認を保存できませんでした。もう一度お試しください。", "error");
  }
}

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  document.getElementById("confirm").addEventListener("click", onConfirm);
  // 宛先・添付の編集に追従して再描画
  item.addHandlerAsync(Office.EventType.RecipientsChanged, render);
  item.addHandlerAsync(Office.EventType.AttachmentsChanged, render);
  render();
});
