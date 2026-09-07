function pay_(userId, shopId, amount, note) {
  return lockRun_(() => {
    const minAmount = Number(getConfig_("MIN_AMOUNT", 1));
    const maxAmount = Number(getConfig_("MAX_AMOUNT", 500));
    amount = Number(amount);

    if (!userId) throw new Error("ユーザーIDが空です。");
    if (!shopId) throw new Error("店IDが空です。");
    if (!Number.isFinite(amount)) throw new Error("金額が不正です。");
    if (amount < minAmount || amount > maxAmount) throw new Error(`金額は ${minAmount}〜${maxAmount} の範囲にしてください。`);

    const user = getUser_(userId);
    if (!user || !user.isActive) throw new Error("ユーザーが見つからない/無効です。");

    const shop = getShop_(shopId);
    if (!shop || !shop.isActive) throw new Error("店が見つからない/無効です。");

    if (user.balance < amount) throw new Error("残高が足りません。");

    const newUserBal = user.balance - amount;
    const newShopBal = shop.balance + amount;

    setUserBalance_(user.userId, newUserBal);
    setShopBalance_(shop.shopId, newShopBal);

    const txId = uuid_();
    const atJst = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");


    appendTx_({
      txId,
      at: atJst, // ★文字列で保存
      type: "PAY",
      userId: user.userId,
      userName: user.name,
      shopId: shop.shopId,
      shopName: shop.shopName,
      amount,
      status: "OK",
      note: note || "",
      meta: JSON.stringify({ mode: "fixedShopQR" }),
    });

    // ★戻り値もDateなし（文字列/数値のみ）
    return {
      ok: true,
      txId,
      at: atJst,
      user: { userId: user.userId, name: user.name, balance: newUserBal },
      shop: { shopId: shop.shopId, shopName: shop.shopName, balance: newShopBal }
    };
  });
}
