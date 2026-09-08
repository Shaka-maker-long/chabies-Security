function optionLetter(index) {
  let n = Number(index) + 1;
  let s = "";
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function normalizeQuotes(quotes) {
  const list = Array.isArray(quotes) ? quotes : [];
  return list.map((q, i) => {
    const copy = q && typeof q === "object" ? Object.assign({}, q) : {};
    const hasOption = String(copy.option || "").trim();
    const hasKind = copy.kind === "option" || copy.kind === "revision";
    copy.option = hasOption ? String(copy.option).trim().toUpperCase() : "A";
    copy.kind = hasKind ? copy.kind : (i === 0 ? "option" : "revision");
    copy.quote_no = String(copy.quote_no || "").trim();
    copy.products = Array.isArray(copy.products) ? copy.products : [];
    return copy;
  });
}

function liveQuoteOptions(quotes) {
  const by = new Map();
  normalizeQuotes(quotes).forEach((q) => {
    if (!q.quote_no && !(q.products && q.products.length)) return;
    by.set(q.option, q);
  });
  return Array.from(by.values()).sort((a, b) => String(a.option).localeCompare(String(b.option)));
}

function quoteOptionsLabel(quotes) {
  return liveQuoteOptions(quotes).map((q) => {
    const no = q.quote_no || "";
    return no ? (no + " " + q.option) : q.option;
  }).filter(Boolean).join(" / ");
}

function nextOptionLetter(quotes) {
  return optionLetter(liveQuoteOptions(quotes).length);
}

function currentOptionLetter(row) {
  const no = String((row && row.quote_no) || "").trim();
  const list = normalizeQuotes(row && row.quotes);
  const hit = list.filter((q) => q.quote_no === no).pop();
  if (hit && hit.option) return hit.option;
  const live = liveQuoteOptions(list);
  return live.length ? live[live.length - 1].option : "A";
}

function findLiveOption(quotes, optionOrNo) {
  const live = liveQuoteOptions(quotes);
  const s = String(optionOrNo || "").trim();
  if (!s) return null;
  const upper = s.toUpperCase();
  return live.find((q) => q.option === upper)
    || live.find((q) => String(q.quote_no || "").toUpperCase() === upper)
    || null;
}

function chosenQuote(enquiry) {
  const list = normalizeQuotes(enquiry && enquiry.quotes);
  const live = liveQuoteOptions(list);
  const no = String((enquiry && enquiry.chosen_quote_no) || "").trim();
  const opt = String((enquiry && enquiry.chosen_option) || "").trim().toUpperCase();
  if (no) {
    const exact = list.filter((q) => q.quote_no === no).pop() || live.find((q) => q.quote_no === no);
    if (exact) return exact;
  }
  if (opt) {
    const hit = live.find((q) => q.option === opt);
    if (hit) return hit;
  }
  if (live.length === 1) return live[0];
  return null;
}

function applyQuoteSnapshot(row, quote) {
  if (!row || !quote) return row;
  if (quote.quote_no) row.quote_no = quote.quote_no;
  if (Array.isArray(quote.products) && quote.products.length) row.products = quote.products;
  if (quote.delivery_incl_vat != null && String(quote.delivery_incl_vat).trim() !== "") {
    row.delivery_incl_vat = quote.delivery_incl_vat;
  }
  if (quote.delivery_excl_vat != null && String(quote.delivery_excl_vat).trim() !== "") {
    row.delivery_excl_vat = quote.delivery_excl_vat;
  }
  row.product = (row.products || []).map((p) => p.product).filter(Boolean).join(", ");
  row.category = (row.products || []).map((p) => p.category).filter(Boolean)[0] || row.category || "";
  return row;
}

function optionSummary(quote) {
  const names = (quote && Array.isArray(quote.products) ? quote.products : [])
    .map((p) => {
      const product = String((p && p.product) || "").trim();
      const variation = String((p && p.variation) || "").trim();
      if (!product) return "";
      return variation ? product + " (" + variation + ")" : product;
    })
    .filter(Boolean);
  return names.join(", ");
}

module.exports = {
  optionLetter,
  normalizeQuotes,
  liveQuoteOptions,
  quoteOptionsLabel,
  nextOptionLetter,
  currentOptionLetter,
  findLiveOption,
  chosenQuote,
  applyQuoteSnapshot,
  optionSummary
};
