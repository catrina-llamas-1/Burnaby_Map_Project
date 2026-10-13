#!/usr/bin/env bash
# One-time setup: lets THIS repo's GitHub Actions deploy to its own,
# dedicated Firebase project (fir-map-d3936) without a stored key
# (Workload Identity Federation). Creates its own service account --
# this project is NOT shared with catrina-llamas-1/Companies-amenities-map.
#
# Run it in Google Cloud Shell (https://shell.cloud.google.com), signed in as an
# owner of the fir-map-d3936 Firebase project:
#   git clone https://github.com/catrina-llamas-1/Burnaby_Map_Project.git map-setup
#   bash ~/map-setup/scripts/setup_github_deploy.sh
#
# At the end it prints two values to add as GitHub repository variables, plus
# a reminder about the GOOGLE_MAPS_API_KEY secret.
# Safe to run again: anything that already exists is reused.

set -euo pipefail

PROJECT_ID="${PROJECT_ID:-fir-map-d3936}"
REPO="${REPO:-catrina-llamas-1/Burnaby_Map_Project}"
SA_NAME="github-deploy"
POOL="github"
PROVIDER="github-repo-burnaby"
SA="${SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"

echo "Project: ${PROJECT_ID}"
echo "Repo:    ${REPO}"
gcloud config set project "${PROJECT_ID}" >/dev/null

echo "== Enabling APIs"
gcloud services enable iam.googleapis.com iamcredentials.googleapis.com sts.googleapis.com \
  firebasehosting.googleapis.com cloudresourcemanager.googleapis.com

# Newly created resources take a little while to become visible/usable
# elsewhere; retry until they are.
wait_for() {
  local what="$1"; shift
  for _ in $(seq 1 30); do
    if "$@" >/dev/null 2>&1; then return 0; fi
    echo "   waiting for ${what}..."
    sleep 5
  done
  echo "ERROR: ${what} still not available after 2.5 minutes; run this script again." >&2
  exit 1
}

echo "== Service account ${SA}"
if ! gcloud iam service-accounts describe "${SA}" >/dev/null 2>&1; then
  gcloud iam service-accounts create "${SA_NAME}" --display-name="GitHub Actions deploy"
fi
wait_for "service account ${SA}" gcloud iam service-accounts describe "${SA}"

echo "== Granting deploy roles"
for role in roles/firebasehosting.admin roles/run.viewer \
            roles/serviceusage.apiKeysViewer roles/serviceusage.serviceUsageConsumer; do
  gcloud projects add-iam-policy-binding "${PROJECT_ID}" \
    --member="serviceAccount:${SA}" --role="${role}" --condition=None >/dev/null
  echo "   ${role}"
done

echo "== Workload identity pool '${POOL}'"
if ! gcloud iam workload-identity-pools describe "${POOL}" --location=global >/dev/null 2>&1; then
  gcloud iam workload-identity-pools create "${POOL}" --location=global --display-name="GitHub Actions"
fi
wait_for "pool '${POOL}'" gcloud iam workload-identity-pools describe "${POOL}" --location=global
PROJECT_NUMBER="$(gcloud projects describe "${PROJECT_ID}" --format='value(projectNumber)')"
POOL_NAME="projects/${PROJECT_NUMBER}/locations/global/workloadIdentityPools/${POOL}"

echo "== OIDC provider '${PROVIDER}' (only ${REPO} can use it)"
if ! gcloud iam workload-identity-pools providers describe "${PROVIDER}" \
      --location=global --workload-identity-pool="${POOL}" >/dev/null 2>&1; then
  gcloud iam workload-identity-pools providers create-oidc "${PROVIDER}" \
    --location=global --workload-identity-pool="${POOL}" \
    --display-name="GitHub Burnaby map" \
    --issuer-uri="https://token.actions.githubusercontent.com" \
    --attribute-mapping="google.subject=assertion.sub,attribute.repository=assertion.repository" \
    --attribute-condition="assertion.repository=='${REPO}'"
fi
wait_for "provider '${PROVIDER}'" gcloud iam workload-identity-pools providers describe "${PROVIDER}" \
  --location=global --workload-identity-pool="${POOL}"

echo "== Letting ${REPO} act as ${SA}"
gcloud iam service-accounts add-iam-policy-binding "${SA}" \
  --role="roles/iam.workloadIdentityUser" \
  --member="principalSet://iam.googleapis.com/${POOL_NAME}/attribute.repository/${REPO}" >/dev/null

cat <<EOF

Done. Now add these two repository VARIABLES in GitHub
(repo -> Settings -> Secrets and variables -> Actions -> Variables tab -> New repository variable):

  Name:  GCP_WORKLOAD_IDENTITY_PROVIDER
  Value: ${POOL_NAME}/providers/${PROVIDER}

  Name:  GCP_SERVICE_ACCOUNT
  Value: ${SA}

Also add a GOOGLE_MAPS_API_KEY secret (Secrets tab, not Variables -- it's
sensitive): a Google Maps Platform API key with the Geocoding API and
Distance Matrix API enabled and billing set up.

EOF
