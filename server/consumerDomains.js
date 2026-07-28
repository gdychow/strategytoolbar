// Task Pane Phase 14: emails on these domains never get a company library,
// not even a "company of one" — reasonably broad, trivially extendable.
const CONSUMER_DOMAINS = new Set([
  "gmail.com",
  "googlemail.com",
  "yahoo.com",
  "ymail.com",
  "hotmail.com",
  "outlook.com",
  "live.com",
  "msn.com",
  "aol.com",
  "icloud.com",
  "me.com",
  "mac.com",
  "protonmail.com",
  "proton.me",
  "gmx.com",
  "gmx.net",
  "mail.com",
  "yandex.com",
  "yandex.ru",
  "zoho.com",
  "fastmail.com",
  "hey.com",
  "qq.com",
  "163.com",
  "naver.com",
]);

/** Lowercases, splits on '@', and returns the email's domain unless it's a consumer provider (or the email is malformed) — in which case null (no company library). */
function deriveCompanyDomain(email) {
  if (typeof email !== "string") return null;
  const at = email.lastIndexOf("@");
  if (at < 0 || at === email.length - 1) return null;
  const domain = email.slice(at + 1).trim().toLowerCase();
  if (!domain || CONSUMER_DOMAINS.has(domain)) return null;
  return domain;
}

module.exports = { CONSUMER_DOMAINS, deriveCompanyDomain };
