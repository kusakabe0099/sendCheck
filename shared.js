/* 宛先・添付の収集、確認済みフラグの保存/取得（commands.js と taskpane.js で共有） */
(function (global) {
  "use strict";

  // 自社ドメインを設定。これ以外は「社外」として扱う。
  const INTERNAL_DOMAINS = ["example.co.jp"];
  const CONFIRMED_KEY = "okanConfirmedSignature";

  function call(fn) {
    return new Promise((resolve, reject) => {
      fn((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(result.error);
      });
    });
  }

  // 短い署名文字列を作る（djb2）。宛先・添付が変わると値が変わる。
  function hash(text) {
    let h = 5381;
    for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0;
    return (h >>> 0).toString(16);
  }

  function domainOf(address) {
    const at = address.lastIndexOf("@");
    return at >= 0 ? address.slice(at + 1).toLowerCase() : "(不明)";
  }

  function isExternal(domain) {
    return !INTERNAL_DOMAINS.includes(domain);
  }

  async function collect(item) {
    const [to, cc, bcc, attachments] = await Promise.all([
      call((cb) => item.to.getAsync(cb)),
      call((cb) => item.cc.getAsync(cb)),
      call((cb) => item.bcc.getAsync(cb)),
      call((cb) => item.getAttachmentsAsync(cb)),
    ]);

    const recipients = [];
    [["To", to], ["Cc", cc], ["Bcc", bcc]].forEach(([kind, list]) => {
      list.forEach((r) =>
        recipients.push({
          kind,
          name: r.displayName || "",
          address: (r.emailAddress || "").toLowerCase(),
        })
      );
    });

    const files = attachments
      .filter((a) => !a.isInline)
      .map((a) => ({ name: a.name, size: a.size }));

    // ドメインごとにまとめる（社外を先頭に）
    const map = new Map();
    recipients.forEach((r) => {
      const domain = domainOf(r.address);
      if (!map.has(domain)) map.set(domain, { domain, external: isExternal(domain), recipients: [] });
      map.get(domain).recipients.push(r);
    });
    const groups = Array.from(map.values()).sort(
      (a, b) => Number(b.external) - Number(a.external) || a.domain.localeCompare(b.domain)
    );

    const signature = hash(
      JSON.stringify([
        recipients.map((r) => r.kind + ":" + r.address).sort(),
        files.map((f) => f.name + ":" + f.size).sort(),
      ])
    );

    return {
      recipients,
      groups,
      files,
      externalCount: recipients.filter((r) => isExternal(domainOf(r.address))).length,
      signature,
    };
  }

  async function getConfirmed(item) {
    try {
      return await call((cb) => item.sessionData.getAsync(CONFIRMED_KEY, cb));
    } catch (e) {
      return null; // 未保存
    }
  }

  function setConfirmed(item, signature) {
    return call((cb) => item.sessionData.setAsync(CONFIRMED_KEY, signature, cb));
  }

  global.Okan = { collect, getConfirmed, setConfirmed };
})(window);
