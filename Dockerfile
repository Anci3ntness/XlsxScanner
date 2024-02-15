FROM node:20

WORKDIR /home/node/app
ADD . .

RUN npm ci

EXPOSE 8080

USER node