FROM node:20-bookworm-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    fonts-noto-core \
    fontconfig \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-guj \
    libjpeg62-turbo \
    libpng16-16 \
    zlib1g \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev
COPY . .
RUN mkdir -p /usr/share/fonts/truetype/custom \
    && (cp -f NotoSansGujarati-Regular.ttf /usr/share/fonts/truetype/custom/ || true) \
    && (cp -f fonts/NotoSansGujarati-Regular.ttf /usr/share/fonts/truetype/custom/ || true) \
    && fc-cache -f || true
RUN mkdir -p /etc/libreoffice/registry /root/.config/libreoffice/4/user \
 && printf '%s\n' \
 '<?xml version="1.0" encoding="UTF-8"?>' \
 '<oor:data xmlns:oor="http://openoffice.org/2001/registry">' \
 ' <dependency file="main"/>' \
 ' <oor:component-data oor:name="Calc" oor:package="org.openoffice.Office">' \
 '  <node oor:name="Formula"><node oor:name="Load">' \
 '   <prop oor:name="OOXMLRecalcMode" oor:type="xs:int"><value>0</value></prop>' \
 '  </node></node>' \
 ' </oor:component-data>' \
 '</oor:data>' > /etc/libreoffice/registry/recalc.xcd \
 && printf '%s\n' \
 '<?xml version="1.0" encoding="UTF-8"?>' \
 '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' \
 '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>' \
 '</oor:items>' > /root/.config/libreoffice/4/user/registrymodifications.xcu
ENV PORT=10000
EXPOSE 10000
CMD ["node", "server.js"]
