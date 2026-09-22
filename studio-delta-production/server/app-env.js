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
  return "STAGING — safe to test updates. Live shop is unchanged. Login Admin / admin";
}

module.exports = {
  envName,
  isStaging,
  appEnv,
  stagingBannerText
};
