"use strict";

function envName() {
  return String(
    process.env.APP_ENV
      || process.env.SD_APP_ENV
      || process.env.RAILWAY_ENVIRONMENT_NAME
      || process.env.RAILWAY_ENVIRONMENT
      || ""
  ).trim().toLowerCase();
}

function isStaging() {
  if (process.env.STAGING === "1" || process.env.STAGING === "true") return true;
  const name = envName();
  return name === "staging" || name === "preview" || name === "updates";
}

function appEnv() {
  if (isStaging()) return "staging";
  const name = envName();
  if (name) return name;
  return process.env.NODE_ENV === "production" ? "production" : "development";
}

function stagingBannerText() {
  // Keep this explicit so the installed PWA makes it obvious updates were committed.
  const tip = String(process.env.SD_STAGING_COMMIT || process.env.RAILWAY_GIT_COMMIT_SHA || "").trim();
  const short = tip ? tip.slice(0, 7) : "";
  const commitBit = short ? (" · committed " + short) : " · updates committed";
  return "STAGING" + commitBit + " — Drawing client approval + quote option recost. Paint PDF + Open PO/QR. Login Admin / admin";
}

module.exports = {
  envName,
  isStaging,
  appEnv,
  stagingBannerText
};
