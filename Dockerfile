# Dockerfile for Collabora Online (coolwsd) development
# This Dockerfile is designed for running coolwsd after setup is complete

FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

# Install runtime dependencies for LibreOffice and coolwsd
# Based on: https://collaboraonline.github.io/post/build-code/#ubuntu
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpng16-16 \
    fontconfig \
    adduser \
    cpio \
    tzdata \
    findutils \
    nano \
    libcap2-bin \
    openssl \
    openssh-client \
    libxcb-shm0 \
    libxcb-render0 \
    libxrender1 \
    libxext6 \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    fonts-droid-fallback \
    fonts-noto-cjk \
    ca-certificates \
    git \
    curl \
    build-essential \
    make \
    automake \
    autoconf \
    libtool \
    pkg-config \
    libsystemd-dev \
    libcap-dev \
    libpng-dev \
    libcppunit-dev \
    libssl-dev \
    libzstd-dev \
    # Poco library - required for coolwsd build
    libpoco-dev \
    python3-polib \
    libpam-dev \
    python3-lxml \
    libgif-dev \
    # Podman - workaround for apparmor_restrict_unprivileged_userns
    podman \
    && rm -rf /var/lib/apt/lists/*

# Install Node.js 20.x (required for browser build - ESLint, etc.)
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && rm -rf /var/lib/apt/lists/*

# Configure git to allow mounted directories (fixes "dubious ownership" error)
RUN git config --global --add safe.directory '*'

# Create the cool user (coolwsd refuses to run as root)
RUN adduser --quiet --system --group --home /opt/cool cool

# Set the working directory
WORKDIR /srv/apps/online

# Copy source, LibreOffice core, and branding into the image (portable build/run)
COPY online/ /srv/apps/online/
COPY libreoffice-mini-25/ /srv/apps/libreoffice-mini-25/
COPY collabora-branding/ /srv/apps/collabora-branding/

# Install browser deps, configure and build
RUN cd /srv/apps/online/browser && npm install
RUN if [ ! -x /srv/apps/online/configure ]; then cd /srv/apps/online && ./autogen.sh; fi
RUN cd /srv/apps/online \
    && ./configure --with-lo-path=/srv/apps/libreoffice-mini-25/instdir \
       --with-lokit-path=/srv/apps/libreoffice-mini-25/include \
       --with-app-branding=/srv/apps/collabora-branding/themes/smartdocs-neutral \
    && make -j $(nproc) \
    && /srv/apps/collabora-branding/install.sh

# Expose the default coolwsd port
EXPOSE 9980


# NOTE: We don't use USER here because:
# - npm install and make need write access to the mounted volume (runs as root)
# - Only coolwsd (make run) needs to run as 'cool' user
# The Makefile's 'run' target already handles switching to cool user internally
# But if it doesn't, we'll handle it in the command

# Default command - runs coolwsd
CMD ["sh", "-c", "set -e; LO_ROOT=${LO_ROOT:-/srv/apps/libreoffice-mini-25/instdir}; cd /srv/apps/online && make run LO_PATH=$LO_ROOT"]
