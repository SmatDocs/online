# Dockerfile for Collabora Online (coolwsd) development
# This Dockerfile is designed for running coolwsd after setup is complete

FROM ubuntu:24.04

ENV DEBIAN_FRONTEND=noninteractive

# Install runtime dependencies for LibreOffice and coolwsd
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
    && rm -rf /var/lib/apt/lists/*

# Set the working directory
WORKDIR /app

# The source code will be mounted at /app via volume
# This allows for development with hot-reload capability

# Expose the default coolwsd port
EXPOSE 9980

# Default command - builds and runs the application
# Uses make run which handles the coolwsd execution
CMD ["sh", "-c", "make -j $(nproc) && make run"]
