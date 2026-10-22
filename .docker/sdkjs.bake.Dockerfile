# ==============================================================================
# MODULE DOCKERFILE
# This file is not meant to be built standalone. It is consumed by the 
# docker-bake.hcl files in the parent monorepos.
#
# REQUIRED CONTEXTS:
# - core-wasm: builds from core-wasm in core repository (not for sdkjs-native)
#
# sdkjs-native: the files the mobile apps' editor engine loads (it draws and
# reads files with core's native code, so it needs no core-wasm)
# ==============================================================================

ARG PRODUCT_VERSION
ARG BUILD_ROOT

#### BASE ####
FROM ubuntu:26.04 AS web-base
    RUN apt-get update && \
        apt-get install -y ca-certificates curl gnupg wget zip brotli bzip2 && \
        curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
        apt-get install -y nodejs && \
        npm install -g @yao-pkg/pkg && \
        rm -rf /var/lib/apt/lists/*

#### SDKJS ####
FROM web-base AS sdkjs-sources

    ARG BUILD_ROOT=/package

    ARG PRODUCT_VERSION

    COPY sdkjs/build/package*.json sdkjs/build/npm-shrinkwrap.json /app/build/

    RUN --mount=type=cache,target=/root/.npm \
        cd app/build && \
        npm ci

    COPY sdkjs/ /app
    COPY sdkjs-forms/ /sdkjs-forms

    ENV BUILD_ROOT=${BUILD_ROOT}

    ENV PRODUCT_VERSION=${PRODUCT_VERSION}

FROM sdkjs-sources AS sdkjs-base

    ## Copy core wasm builds
    COPY --from=core-wasm ${BUILD_ROOT}/engine/ /app/pdf/src/engine/
    COPY --from=core-wasm ${BUILD_ROOT}/zlib/ /app/common/zlib/
    COPY --from=core-wasm ${BUILD_ROOT}/hash/ /app/common/hash/
    COPY --from=core-wasm ${BUILD_ROOT}/spell/ /app/common/spell/
    COPY --from=core-wasm ${BUILD_ROOT}/libfont/ /app/common/libfont/

FROM sdkjs-base AS sdkjs-desktop
    ENV SDK_ADDONS=/sdkjs-forms
    ENV SDK_PLATFORM=desktop
    RUN cd app/build && \
        npm run build

FROM sdkjs-base AS sdkjs
    ENV SDK_ADDONS=/sdkjs-forms
    RUN cd app/build && \
        npm run build

#### SDKJS FOR THE MOBILE EDITORS' ENGINE ####
FROM sdkjs-sources AS sdkjs-native-build
    ENV SDK_ADDONS=/sdkjs-forms
    ENV SKIP_DEVELOP=1
    RUN cd app/build && \
        npm run build

    # What the engine loads: the bundles, sdkjs' native glue, presentation designs
    # (also x2t's theme directory when saving), SmartArt definitions and chart
    # styles (the web editor loads these over the web), the third-party list for
    # the apps' licences page and xregexp (DoctRenderer.config)
    RUN <<EOF
        set -e
        DEPLOY=${BUILD_ROOT}/sdkjs
        OUT=/out
        for editor in word slide cell; do
            mkdir -p $OUT/$editor
            cp $DEPLOY/$editor/sdk-all-min.js $DEPLOY/$editor/sdk-all.js $OUT/$editor/
        done
        cp -r $DEPLOY/slide/themes $OUT/slide/
        mkdir -p $OUT/common/Native $OUT/common/libfont/engine $OUT/pdf/src/engine $OUT/common/SmartArts $OUT/common/Charts $OUT/vendor
        cp $DEPLOY/common/Native/*.js $OUT/common/Native/
        cp $DEPLOY/common/libfont/engine/fonts_native.js $OUT/common/libfont/engine/
        cp $DEPLOY/pdf/src/engine/drawingfile_native.js $OUT/pdf/src/engine/
        cp -r /app/common/SmartArts/SmartArtDrawing /app/common/SmartArts/SmartArtData $OUT/common/SmartArts/
        cp /app/common/Charts/ChartStyles.js $OUT/common/Charts/
        cp /app/3DPARTY.md $OUT/
        cp /app/vendor/xregexp-all-min.js $OUT/vendor/
EOF

FROM scratch AS sdkjs-native
    COPY --from=sdkjs-native-build /out/ /
