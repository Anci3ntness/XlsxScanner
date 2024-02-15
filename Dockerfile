FROM node:20

WORKDIR /home/node/app
ADD . .

RUN npm ci

USER node