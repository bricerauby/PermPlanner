FROM ubuntu:18.04
RUN apt-get update && apt -y install locales
RUN locale-gen en_US.UTF-8
ENV LANG en_US.UTF-8
ENV LANGUAGE en_US:en
ENV LC_ALL en_US.UTF-8
 
RUN apt-get -qq -y install curl
RUN apt install -y  build-essential python3-dev python3-pip python3-setuptools python3-wheel python3-cffi libcairo2 libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0 libffi-dev shared-mime-info zip

RUN pip3 install weasyprint openpyxl tqdm
ADD  ./*.py /app/
ADD ./perms /app/perms
ADD ./logo.png /app/
ADD ./Liste_Staffeurs.xlsx /app/
WORKDIR /app 
