/* 送信時イベント（OnMessageSend）のハンドラ */
Office.onReady();

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const state = await Okan.collect(item);

    // 宛先なしはOutlook側のエラーに任せる
    if (state.recipients.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    // 確認画面で承認済み、かつ宛先・添付が変わっていなければ送信を許可
    const confirmed = await Okan.getConfirmed(item);
    if (confirmed === state.signature) {
      event.completed({ allowEvent: true });
      return;
    }

    event.completed({
      allowEvent: false,
      errorMessage:
        "送信前に宛先と添付ファイルの確認が必要です。" +
        `（宛先 ${state.recipients.length} 件、うち社外 ${state.externalCount} 件、添付 ${state.files.length} 件）`,
      cancelLabel: "確認画面を開く", // 「アクションを実行」ボタンのラベル（20文字以内）
      commandId: "msgComposeOpenPaneButton", // manifest.xml のボタンIDと一致させる
    });
  } catch (e) {
    event.completed({
      allowEvent: false,
      errorMessage: "送信前チェックでエラーが発生しました。もう一度「送信」を押してください。",
    });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
