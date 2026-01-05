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

# Set the working directory (must match host path since Makefile has absolute paths)
WORKDIR /srv/apps/online

# The source code will be mounted via volume
# This allows for development with hot-reload capability

# Expose the default coolwsd port
EXPOSE 9980

# Default command - installs npm deps, builds and runs the application
CMD ["sh", "-c", "cd /srv/apps/online/browser && npm install && cd /srv/apps/online && make -j $(nproc) && make run"]

