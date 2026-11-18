/**
 * SPDX-FileCopyrightText: 2026 Euro-Office contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */

/**
 * Shared env-var parsing used identically by webpack.sdk.factory.mjs,
 * scripts/build-develop.cjs, scripts/build-pipeline.cjs, and
 * scripts/deploy-assets.cjs — extracted so these copies can't silently
 * diverge (see build/lib/sdk-configs.cjs for the same rationale applied
 * to config loading).
 */

'use strict';

const path = require('path');
const fs   = require('fs');

// process.env.SDK_ADDONS: path.delimiter-separated list of addon directories.
function parseAddonDirs(env) {
    env = env || process.env;
    return env.SDK_ADDONS
        ? env.SDK_ADDONS.split(path.delimiter).filter(Boolean)
        : [];
}

// process.env.BUILD_ROOT, resolved to the sdkjs-specific deploy dir.
// buildDir is the caller's build/ directory (path.resolve(__dirname, ...)),
// since the default falls back to <sdkjs-root>/deploy/sdkjs.
function resolveBuildRoot(buildDir, env) {
    env = env || process.env;
    return env.BUILD_ROOT
        ? path.resolve(env.BUILD_ROOT, 'sdkjs')
        : path.resolve(buildDir, '..', 'deploy', 'sdkjs');
}

// Defaults for APP_COPYRIGHT/PUBLISHER_URL — shared with webpack.sdk.factory.mjs's
// cache.version salt so the two copies of these strings can't drift apart.
function defaultAppCopyright() {
    return `Copyright (C) Ascensio System SIA 2012-2025. All rights reserved; Euro-Office contributors 2026 - ${new Date().getFullYear()}`;
}
const DEFAULT_PUBLISHER_URL = 'https://github.com/Euro-Office/';

// Reads build/license.header and substitutes the @@AppCopyright/@@PublisherUrl/
// @@Version/@@Build placeholders. Shared by webpack.sdk.factory.mjs (banner
// injected via BannerPlugin after minification) and deploy-assets.cjs (banner
// prepended directly to each copied JS file) so a future change to a placeholder
// or the default copyright string can't be applied to only one copy.
//
// buildDir is the caller's build/ directory (where license.header lives).
function buildLicenseHeader(buildDir, { env } = {}) {
    env = env || process.env;

    const appCopyright = env.APP_COPYRIGHT || defaultAppCopyright();
    const publisherUrl = env.PUBLISHER_URL || DEFAULT_PUBLISHER_URL;
    const version       = env.PRODUCT_VERSION || '0.0.0';
    const buildNumber   = env.BUILD_NUMBER     || '0';

    let licenseText = fs.readFileSync(path.join(buildDir, 'license.header'), 'utf8');
    licenseText = licenseText
        .replace('@@AppCopyright', appCopyright)
        .replace('@@PublisherUrl', publisherUrl)
        .replace('@@Version', version)
        .replace('@@Build', buildNumber);

    return licenseText;
}

module.exports = { parseAddonDirs, resolveBuildRoot, buildLicenseHeader, defaultAppCopyright, DEFAULT_PUBLISHER_URL };
