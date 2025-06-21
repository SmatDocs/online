# Use a standard Ubuntu base image
FROM ubuntu:24.04

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# Update and install all build dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    python3 \
    python3-lxml \
    git \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app