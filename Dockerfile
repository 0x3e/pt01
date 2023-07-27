FROM ubuntu:22.04
RUN echo 'APT::Install-Suggests "0";' >> /etc/apt/apt.conf.d/00-docker
RUN echo 'APT::Install-Recommends "0";' >> /etc/apt/apt.conf.d/00-docker
RUN DEBIAN_FRONTEND=noninteractive \
  apt-get update \
  && apt-get install -y python3 \
  && rm -rf /var/lib/apt/lists/*
#wine
RUN apt-get update && apt-get install -y wget software-properties-common gnupg winbind xvfb
RUN dpkg --add-architecture i386 && \
    wget -nc https://dl.winehq.org/wine-builds/winehq.key && \
    apt-key add winehq.key && \
    add-apt-repository 'deb https://dl.winehq.org/wine-builds/ubuntu/ jammy main' && \
    apt-get update && apt-get install -y winehq-stable winetricks

ENV WINEDEBUG=fixme-all

RUN apt-get update && apt-get install -y screen

RUN useradd -ms /bin/bash apprunner
USER apprunner
WORKDIR /app
ENTRYPOINT ["/app/start_screen.sh"]
