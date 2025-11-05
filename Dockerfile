# Utiliser une image Nginx officielle
FROM nginx:alpine

# Copier le contenu de ton projet dans le conteneur
COPY . /usr/share/nginx/html

# Copier la configuration Nginx personnalisée (optionnel)
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Exposer le port 80
EXPOSE 80

# Démarrer Nginx
CMD ["nginx", "-g", "daemon off;"]
