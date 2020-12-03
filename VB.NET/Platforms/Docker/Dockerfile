FROM mcr.microsoft.com/dotnet/core/runtime:3.1-buster-slim AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/core/sdk:3.1-buster AS build
WORKDIR /src
COPY ["SpreadsheetDocker.vbproj", ""]
RUN dotnet restore "./SpreadsheetDocker.vbproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "SpreadsheetDocker.vbproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "SpreadsheetDocker.vbproj" -c Release -o /app/publish

FROM base AS final

# Update package sources to include supplemental packages (contrib archive area).
RUN sed -i 's/main/main contrib/g' /etc/apt/sources.list

# Downloads the package lists from the repositories.
RUN apt-get update

# Install System.Drawing.Common dependency.
RUN apt-get install -y libgdiplus

# Install Microsoft TrueType core fonts.
RUN apt-get install -y ttf-mscorefonts-installer

# Or install Liberation TrueType fonts.
# RUN apt-get install -y fonts-liberation

# Or some other font package...

WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "SpreadsheetDocker.dll"]