# Use the official .NET runtime as a base image
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 8080

# Use the SDK image to build the app
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["word_2_pdf.csproj", "./"]
RUN dotnet restore "./word_2_pdf.csproj"
COPY . .
RUN dotnet build "word_2_pdf.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "word_2_pdf.csproj" -c Release -o /app/publish

# Final stage: copy the build and run
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
# Ensure we listen on the port Render provides
ENTRYPOINT ["dotnet", "word_2_pdf.dll"]
