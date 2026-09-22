#!/usr/bin/env node
"use strict";

/**
 * Create a Railway staging environment for Studio Delta updates.
 *
 * Needs RAILWAY_TOKEN (account or workspace token) in the environment:
 *   https://railway.com/account/tokens
 *
 * Usage:
 *   RAILWAY_TOKEN=... node scripts/create-railway-staging.js
 *
 * Project defaults match the live zucchini-rebirth production deploy.
 */

const PROJECT_ID = process.env.RAILWAY_PROJECT_ID || "b06f4c9a-555d-4c71-a2d0-e3f246dbc16c";
const SOURCE_ENV_ID = process.env.RAILWAY_SOURCE_ENV_ID || "ea403c39-9992-4548-b6c3-01a42208388e";
const STAGING_NAME = process.env.RAILWAY_STAGING_NAME || "staging";
const STAGING_BRANCH = process.env.RAILWAY_STAGING_BRANCH || "cursor/staging-updates-9723";
const API = "https://backboard.railway.com/graphql/v2";

async function gql(token, query, variables) {
  const res = await fetch(API, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ query, variables })
  });
  const json = await res.json();
  if (!res.ok || (json.errors && json.errors.length)) {
    const msg = json.errors ? json.errors.map((e) => e.message).join("; ") : res.statusText;
    throw new Error(msg || "Railway API error");
  }
  return json.data;
}

async function main() {
  const token = String(process.env.RAILWAY_TOKEN || "").trim();
  if (!token) {
    console.error("Set RAILWAY_TOKEN (https://railway.com/account/tokens) then re-run.");
    process.exit(2);
  }

  const existing = await gql(token, `query($projectId: String!) {
    environments(projectId: $projectId) {
      edges { node { id name } }
    }
  }`, { projectId: PROJECT_ID });

  const edges = (((existing || {}).environments || {}).edges) || [];
  let staging = edges.map((e) => e.node).find((n) => String(n.name).toLowerCase() === STAGING_NAME.toLowerCase());

  if (!staging) {
    console.log("Creating environment", STAGING_NAME, "from", SOURCE_ENV_ID);
    const created = await gql(token, `mutation($input: EnvironmentCreateInput!) {
      environmentCreate(input: $input) { id name }
    }`, {
      input: {
        projectId: PROJECT_ID,
        name: STAGING_NAME,
        sourceEnvironmentId: SOURCE_ENV_ID
      }
    });
    staging = created.environmentCreate;
    console.log("Created", staging.id, staging.name);
  } else {
    console.log("Staging already exists", staging.id, staging.name);
  }

  const envDetail = await gql(token, `query($id: String!) {
    environment(id: $id) {
      id
      name
      serviceInstances {
        edges {
          node {
            id
            serviceId
            serviceName
            domains { serviceDomains { domain } }
          }
        }
      }
    }
  }`, { id: staging.id });

  const instances = ((((envDetail || {}).environment || {}).serviceInstances || {}).edges) || [];
  for (const edge of instances) {
    const node = edge.node || {};
    const serviceId = node.serviceId;
    if (!serviceId) continue;
    console.log("Configuring service", node.serviceName || serviceId);

    await gql(token, `mutation($serviceId: String!, $environmentId: String!, $input: ServiceInstanceUpdateInput!) {
      serviceInstanceUpdate(serviceId: $serviceId, environmentId: $environmentId, input: $input)
    }`, {
      serviceId,
      environmentId: staging.id,
      input: {
        branch: STAGING_BRANCH
      }
    }).catch((err) => {
      console.warn("branch update skipped:", err.message);
    });

    for (const [name, value] of [
      ["APP_ENV", "staging"],
      ["STAGING", "1"],
      ["TZ", "Africa/Johannesburg"],
      ["DATA_DIR", "/app/data"],
      ["LOCAL_ADMIN_CODE", "admin"]
    ]) {
      await gql(token, `mutation($input: VariableUpsertInput!) {
        variableUpsert(input: $input)
      }`, {
        input: {
          projectId: PROJECT_ID,
          environmentId: staging.id,
          serviceId,
          name,
          value
        }
      }).catch((err) => {
        console.warn("variable", name, err.message);
      });
    }

    try {
      await gql(token, `mutation($input: VolumeCreateInput!) {
        volumeCreate(input: $input) { id name }
      }`, {
        input: {
          projectId: PROJECT_ID,
          environmentId: staging.id,
          serviceId,
          mountPath: "/app/data"
        }
      });
      console.log("Volume /app/data attached (or already present)");
    } catch (err) {
      console.warn("volume:", err.message);
    }

    try {
      const domain = await gql(token, `mutation($input: ServiceDomainCreateInput!) {
        serviceDomainCreate(input: $input) { id domain }
      }`, {
        input: {
          environmentId: staging.id,
          serviceId
        }
      });
      console.log("Domain", domain.serviceDomainCreate && domain.serviceDomainCreate.domain);
    } catch (err) {
      console.warn("domain:", err.message);
    }

    try {
      const dep = await gql(token, `mutation($serviceId: String!, $environmentId: String!) {
        serviceInstanceDeployV2(serviceId: $serviceId, environmentId: $environmentId)
      }`, { serviceId, environmentId: staging.id });
      console.log("Deploy triggered", dep.serviceInstanceDeployV2);
    } catch (err) {
      console.warn("deploy:", err.message);
    }
  }

  console.log("\nStaging environment id:", staging.id);
  console.log("Open: https://railway.com/project/" + PROJECT_ID + "?environmentId=" + staging.id);
  console.log("Push branch", STAGING_BRANCH, "to GitHub to update staging without touching live.");
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
