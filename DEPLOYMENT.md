# SmartDraft Add-in Deployment Guide

## Overview

SmartDraft is a professional AI-powered email assistant add-in for Microsoft Outlook that works on both **Outlook Web** and **Outlook Desktop** applications. This guide provides detailed instructions for deploying and configuring the add-in.

## Prerequisites

- Node.js 16.x or higher
- npm or yarn package manager
- Microsoft Office 365 subscription or Outlook Web Access
- Administrative access to Microsoft 365 tenant (for deployment)
- HTTPS-enabled web server for hosting the add-in files

## Installation Steps

### 1. Clone the Repository

```bash
git clone https://github.com/khanshahnawaz185/smartdraft-addin.git
cd smartdraft-addin
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Build the Project

```bash
npm run build
```

### 4. Generate Development Certificates (for HTTPS)

```bash
npm run generate-certs
```

## Deployment to Outlook Web

### Option 1: Centralized Deployment (Recommended)

1. **Upload to Microsoft 365 Admin Center**
   - Go to https://admin.microsoft.com
   - Navigate to Settings > Integrated Apps
   - Click "Upload custom apps"
   - Select `manifest.xml` from the root directory
   - Click "Deploy"
   - Choose users/groups to deploy to
   - Complete the deployment process

2. **Access in Outlook Web**
   - Open Outlook at https://outlook.office.com
   - The SmartDraft add-in will appear in the ribbon
   - Click "SmartDraft" to open the AI Assistant task pane

### Option 2: Manual Installation (Testing)

1. **Host the Files on HTTPS Server**
   - Upload all files from `src/taskpane/` to your web server
   - Ensure HTTPS is enabled
   - Note the URLs (e.g., https://yourdomain.com/smartdraft/)

2. **Update Manifest File**
   - Edit `manifest.xml`
   - Replace placeholder URLs with your server URLs:
     ```xml
     <bt:String id="functionfile" DefaultValue="https://yourdomain.com/smartdraft/functions.html"/>
     <bt:String id="taskpaneurl" DefaultValue="https://yourdomain.com/smartdraft/taskpane.html"/>
     ```

3. **Install via Outlook Web**
   - In Outlook Web, click the "..." menu
   - Select "Get Add-ins"
   - Click "My add-ins" > "Upload My Add-in"
   - Upload the updated `manifest.xml`
   - Click "Install"

## Deployment to Outlook Desktop

### Windows Outlook

1. **Configure Manifest URL**
   - Update the manifest file with your server URLs (same as above)
   - Host the manifest.xml on your HTTPS server

2. **Add the Manifest in Outlook**
   - Open Outlook Desktop
   - Go to File > Options > Trust Center > Trust Center Settings
   - Click "Trusted Add-in Catalogs"
   - Add your manifest URL to "Manifest URL"
   - Click OK
   - Restart Outlook
   - The add-in will appear in your ribbon

### Mac Outlook

1. **Network Configuration**
   - Ensure your add-in is accessible via HTTPS
   - Update manifest with correct URLs

2. **Installation**
   - Open Outlook for Mac
   - Click "Insert" menu
   - Select "Get Add-ins"
   - Click "My Add-ins" > "Upload My Add-in"
   - Upload manifest.xml
   - Click "Install"

## Configuration

### Updating Manifest.xml

The manifest.xml file contains essential configuration. Key elements:

```xml
<!-- Add-in ID (generate unique GUID) -->
<Id>YOUR-UNIQUE-ID</Id>

<!-- Version number -->
<Version>1.0.0.0</Version>

<!-- Display name -->
<DisplayName DefaultValue="SmartDraft - AI Email Assistant"/>

<!-- Resource URLs -->
<bt:String id="taskpaneurl" DefaultValue="https://yourdomain.com/taskpane.html"/>
<bt:String id="functionfile" DefaultValue="https://yourdomain.com/functions.html"/>
```

### Environment Variables

Create a `.env` file in the root directory:

```
API_KEY=your_openai_api_key
API_ENDPOINT=https://api.openai.com/v1
OUTLOOK_API_VERSION=1.12
```

## Features

### Email Analysis
- **Sentiment Detection**: Analyzes the emotional tone of incoming emails
- **Urgency Detection**: Identifies high-priority messages
- **Key Points Extraction**: Summarizes main topics

### Smart Reply Generation
- Context-aware response suggestions
- Professional tone maintenance
- Multiple response templates

## Troubleshooting

### Add-in Not Appearing

1. **Check HTTPS Configuration**
   - Ensure your server uses valid SSL/TLS certificate
   - Verify manifest URLs are accessible

2. **Clear Cache**
   - Outlook Web: Clear browser cache and cookies
   - Outlook Desktop: Restart application

3. **Validate Manifest**
   ```bash
   npm run validate
   ```

### Permission Errors

- Ensure manifest has `<Permissions>ReadWriteItem</Permissions>`
- User must grant permissions when first using the add-in

### Network Issues

- Verify CORS headers are properly configured
- Check firewall rules for HTTPS traffic on port 443

## Testing

### Local Testing

```bash
npm start
```

This starts a development server on `https://localhost:3000`

### Manual Testing Checklist

- [ ] Add-in loads in Outlook Web
- [ ] Add-in loads in Outlook Desktop
- [ ] "Analyze Email" button works
- [ ] "Generate Reply" button works
- [ ] Sentiment analysis displays correctly
- [ ] Urgency levels are accurate
- [ ] Key points are extracted

## Security

- All communication uses HTTPS
- Manifest is signed and validated
- User data is processed locally when possible
- API calls use secure authentication
- No email content is stored on servers

## Support

For issues or questions:
- Check GitHub Issues: https://github.com/khanshahnawaz185/smartdraft-addin/issues
- Review manifest validation errors
- Enable debug mode for detailed logging

## Updates

To update the add-in:

1. Update version in `package.json` and `manifest.xml`
2. Run `npm run build`
3. Deploy new files to your hosting server
4. Users will receive update prompts in Outlook

## License

MIT License - See LICENSE file for details
