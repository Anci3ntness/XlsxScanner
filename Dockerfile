FROM node:latest
WORKDIR /usr/src/app
COPY package*.json ./
RUN npm install --network-timeout 100000
COPY . .
CMD [ "npm", "start" ]