# Usar una imagen base de .NET SDK para compilar
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build

# Configurar el directorio de trabajo
WORKDIR /app

# Copiar el archivo de proyecto y restaurar las dependencias
COPY *.csproj .
RUN dotnet restore

# Copiar todo el resto del c�digo y construir
COPY . .
RUN dotnet publish -c Release -o out

# Usar una imagen base de .NET Runtime para ejecutar la aplicaci�n
FROM mcr.microsoft.com/dotnet/aspnet:6.0

# Configurar el directorio de trabajo en el contenedor
WORKDIR /app

# Copiar los archivos publicados desde el contenedor de construcci�n
COPY --from=build /app/out .

# Especificar el comando para ejecutar la aplicaci�n
ENTRYPOINT ["dotnet", "GestorDelal.dll"]
