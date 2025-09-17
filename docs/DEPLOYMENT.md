# Excel GPT Middleware Deployment Guide

This guide covers deployment options, production configuration, monitoring, and maintenance for the Excel GPT Middleware.

## Table of Contents

1. [Production Environment Setup](#production-environment-setup)
2. [Docker Deployment](#docker-deployment)
3. [Cloud Deployment](#cloud-deployment)
4. [Security Hardening](#security-hardening)
5. [Monitoring and Logging](#monitoring-and-logging)
6. [Backup and Recovery](#backup-and-recovery)
7. [Maintenance](#maintenance)
8. [Troubleshooting](#troubleshooting)

## Production Environment Setup

### System Requirements

**Minimum Requirements:**
- CPU: 2 cores
- RAM: 4GB
- Storage: 20GB
- Network: Stable internet connection
- OS: Windows Server 2019+, Ubuntu 18.04+, or CentOS 7+

**Recommended Requirements:**
- CPU: 4+ cores
- RAM: 8GB+
- Storage: 50GB+ SSD
- Network: High-speed internet with redundancy
- Load balancer for high availability

### Environment Configuration

Create production environment file:

```bash
# Copy and customize for production
cp .env.example .env.production
```

**Production .env.production:**
```env
# Environment
NODE_ENV=production
PORT=3000

# Azure AD Configuration
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret

# Security
JWT_SECRET=your-super-secure-jwt-secret-64-chars-minimum
API_KEY=your-production-api-key

# Logging
LOG_LEVEL=info
LOG_DIR=/var/log/excel-gpt-middleware

# Rate Limiting (Production values)
RATE_LIMIT_WINDOW_MS=900000
RATE_LIMIT_MAX_REQUESTS=1000

# CORS (Restrict to your domains)
ALLOWED_ORIGINS=https://your-gpt-domain.com,https://your-app-domain.com

# Trust proxy (if behind load balancer)
TRUST_PROXY=true

# Performance
MAX_REQUEST_SIZE=50mb
REQUEST_TIMEOUT=60000
```

### SSL/TLS Configuration

**Option 1: Reverse Proxy (Recommended)**
Use nginx or Apache as a reverse proxy:

```nginx
# /etc/nginx/sites-available/excel-gpt-middleware
server {
    listen 443 ssl http2;
    server_name your-domain.com;

    ssl_certificate /path/to/your/certificate.crt;
    ssl_certificate_key /path/to/your/private.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-RSA-AES256-GCM-SHA512:DHE-RSA-AES256-GCM-SHA512;

    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_cache_bypass $http_upgrade;
    }
}
```

**Option 2: Application-Level SSL**
Install SSL certificates and configure Express:

```javascript
// In server.js
const https = require('https');
const fs = require('fs');

const options = {
    key: fs.readFileSync('/path/to/private.key'),
    cert: fs.readFileSync('/path/to/certificate.crt')
};

https.createServer(options, app).listen(443, () => {
    console.log('HTTPS Server running on port 443');
});
```

## Docker Deployment

### Dockerfile

```dockerfile
# Use official Node.js runtime
FROM node:18-alpine

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production

# Copy application code
COPY src/ ./src/
COPY docs/ ./docs/
COPY examples/ ./examples/

# Create logs directory
RUN mkdir -p /app/logs

# Create non-root user
RUN addgroup -g 1001 -S nodejs
RUN adduser -S nodejs -u 1001

# Change ownership
RUN chown -R nodejs:nodejs /app
USER nodejs

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD node -e "require('http').get('http://localhost:3000/health', (res) => { process.exit(res.statusCode === 200 ? 0 : 1) })"

# Start application
CMD ["npm", "start"]
```

### Docker Compose

```yaml
# docker-compose.yml
version: '3.8'

services:
  excel-gpt-middleware:
    build: .
    ports:
      - "3000:3000"
    environment:
      - NODE_ENV=production
    env_file:
      - .env.production
    volumes:
      - ./logs:/app/logs
      - ./config:/app/config
    restart: unless-stopped
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:3000/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
    networks:
      - excel-middleware-network

  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf
      - ./ssl:/etc/nginx/ssl
    depends_on:
      - excel-gpt-middleware
    restart: unless-stopped
    networks:
      - excel-middleware-network

networks:
  excel-middleware-network:
    driver: bridge
```

### Build and Deploy

```bash
# Build Docker image
docker build -t excel-gpt-middleware .

# Run with Docker Compose
docker-compose up -d

# Check status
docker-compose ps

# View logs
docker-compose logs -f excel-gpt-middleware
```

## Cloud Deployment

### Azure Container Instances

```bash
# Create resource group
az group create --name excel-middleware-rg --location eastus

# Create container instance
az container create \
  --resource-group excel-middleware-rg \
  --name excel-gpt-middleware \
  --image your-registry/excel-gpt-middleware:latest \
  --cpu 2 \
  --memory 4 \
  --ports 3000 \
  --environment-variables NODE_ENV=production \
  --secure-environment-variables \
    AZURE_CLIENT_SECRET=your-secret \
    JWT_SECRET=your-jwt-secret
```

### Azure App Service

```bash
# Create App Service plan
az appservice plan create \
  --name excel-middleware-plan \
  --resource-group excel-middleware-rg \
  --sku B1 \
  --is-linux

# Create web app
az webapp create \
  --resource-group excel-middleware-rg \
  --plan excel-middleware-plan \
  --name excel-gpt-middleware-app \
  --deployment-container-image-name your-registry/excel-gpt-middleware:latest

# Configure environment variables
az webapp config appsettings set \
  --resource-group excel-middleware-rg \
  --name excel-gpt-middleware-app \
  --settings NODE_ENV=production
```

### AWS ECS

```yaml
# task-definition.json
{
  "family": "excel-gpt-middleware",
  "networkMode": "awsvpc",
  "requiresCompatibilities": ["FARGATE"],
  "cpu": "512",
  "memory": "1024",
  "executionRoleArn": "arn:aws:iam::account:role/ecsTaskExecutionRole",
  "containerDefinitions": [
    {
      "name": "excel-gpt-middleware",
      "image": "your-registry/excel-gpt-middleware:latest",
      "portMappings": [
        {
          "containerPort": 3000,
          "protocol": "tcp"
        }
      ],
      "environment": [
        {
          "name": "NODE_ENV",
          "value": "production"
        }
      ],
      "secrets": [
        {
          "name": "AZURE_CLIENT_SECRET",
          "valueFrom": "arn:aws:secretsmanager:region:account:secret:excel-middleware-secrets"
        }
      ],
      "logConfiguration": {
        "logDriver": "awslogs",
        "options": {
          "awslogs-group": "/ecs/excel-gpt-middleware",
          "awslogs-region": "us-east-1",
          "awslogs-stream-prefix": "ecs"
        }
      }
    }
  ]
}
```

### Google Cloud Run

```bash
# Build and push to Google Container Registry
gcloud builds submit --tag gcr.io/your-project/excel-gpt-middleware

# Deploy to Cloud Run
gcloud run deploy excel-gpt-middleware \
  --image gcr.io/your-project/excel-gpt-middleware \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --set-env-vars NODE_ENV=production \
  --set-secrets AZURE_CLIENT_SECRET=excel-middleware-secrets:latest
```

## Security Hardening

### Application Security

1. **Environment Variables**
```bash
# Use secure secret management
# Never commit secrets to version control
# Rotate secrets regularly
# Use different secrets for each environment
```

2. **Network Security**
```nginx
# Rate limiting at proxy level
limit_req_zone $binary_remote_addr zone=api:10m rate=10r/s;
limit_req zone=api burst=20 nodelay;

# Block suspicious requests
if ($request_method !~ ^(GET|POST|HEAD|OPTIONS)$ ) {
    return 405;
}

# Security headers
add_header X-Frame-Options DENY;
add_header X-Content-Type-Options nosniff;
add_header X-XSS-Protection "1; mode=block";
add_header Strict-Transport-Security "max-age=31536000; includeSubDomains";
```

3. **Application Hardening**
```javascript
// Additional security middleware
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            scriptSrc: ["'self'"],
            styleSrc: ["'self'", "'unsafe-inline'"],
            imgSrc: ["'self'", "data:", "https:"],
        },
    },
    hsts: {
        maxAge: 31536000,
        includeSubDomains: true,
        preload: true
    }
}));
```

### Infrastructure Security

1. **Firewall Rules**
```bash
# Allow only necessary ports
ufw allow 22/tcp    # SSH
ufw allow 80/tcp    # HTTP
ufw allow 443/tcp   # HTTPS
ufw deny 3000/tcp   # Block direct access to app
ufw enable
```

2. **User Permissions**
```bash
# Create dedicated user
useradd -r -s /bin/false excel-middleware
chown -R excel-middleware:excel-middleware /opt/excel-gpt-middleware
```

3. **Log Security**
```bash
# Secure log files
chmod 640 /var/log/excel-gpt-middleware/*.log
chown excel-middleware:adm /var/log/excel-gpt-middleware/*.log
```

## Monitoring and Logging

### Application Monitoring

**Health Check Monitoring:**
```bash
# Create monitoring script
#!/bin/bash
# health-check.sh
HEALTH_URL="https://your-domain.com/health/detailed"
RESPONSE=$(curl -s -o /dev/null -w "%{http_code}" $HEALTH_URL)

if [ $RESPONSE -eq 200 ]; then
    echo "Service is healthy"
    exit 0
else
    echo "Service is unhealthy (HTTP $RESPONSE)"
    exit 1
fi
```

**Log Aggregation:**
```yaml
# docker-compose.logging.yml
version: '3.8'

services:
  elasticsearch:
    image: docker.elastic.co/elasticsearch/elasticsearch:7.15.0
    environment:
      - discovery.type=single-node
    volumes:
      - elasticsearch-data:/usr/share/elasticsearch/data

  logstash:
    image: docker.elastic.co/logstash/logstash:7.15.0
    volumes:
      - ./logstash.conf:/usr/share/logstash/pipeline/logstash.conf
      - ./logs:/logs

  kibana:
    image: docker.elastic.co/kibana/kibana:7.15.0
    ports:
      - "5601:5601"
    environment:
      - ELASTICSEARCH_HOSTS=http://elasticsearch:9200

volumes:
  elasticsearch-data:
```

### System Monitoring

**Prometheus Configuration:**
```yaml
# prometheus.yml
global:
  scrape_interval: 15s

scrape_configs:
  - job_name: 'excel-gpt-middleware'
    static_configs:
      - targets: ['localhost:3000']
    metrics_path: '/metrics'
    scrape_interval: 30s
```

**Grafana Dashboard:**
```json
{
  "dashboard": {
    "title": "Excel GPT Middleware",
    "panels": [
      {
        "title": "Request Rate",
        "type": "graph",
        "targets": [
          {
            "expr": "rate(http_requests_total[5m])"
          }
        ]
      },
      {
        "title": "Response Time",
        "type": "graph",
        "targets": [
          {
            "expr": "histogram_quantile(0.95, rate(http_request_duration_seconds_bucket[5m]))"
          }
        ]
      }
    ]
  }
}
```

### Alerting

**Alert Rules:**
```yaml
# alert-rules.yml
groups:
  - name: excel-middleware-alerts
    rules:
      - alert: HighErrorRate
        expr: rate(http_requests_total{status=~"5.."}[5m]) > 0.1
        for: 5m
        labels:
          severity: critical
        annotations:
          summary: "High error rate detected"

      - alert: ServiceDown
        expr: up{job="excel-gpt-middleware"} == 0
        for: 1m
        labels:
          severity: critical
        annotations:
          summary: "Excel GPT Middleware is down"
```

## Backup and Recovery

### Data Backup

```bash
#!/bin/bash
# backup.sh
BACKUP_DIR="/backup/excel-middleware"
DATE=$(date +%Y%m%d_%H%M%S)

# Create backup directory
mkdir -p $BACKUP_DIR

# Backup configuration
tar -czf $BACKUP_DIR/config_$DATE.tar.gz /opt/excel-gpt-middleware/config

# Backup logs (last 30 days)
find /var/log/excel-gpt-middleware -name "*.log" -mtime -30 -exec tar -czf $BACKUP_DIR/logs_$DATE.tar.gz {} +

# Cleanup old backups (keep 7 days)
find $BACKUP_DIR -name "*.tar.gz" -mtime +7 -delete
```

### Disaster Recovery

**Recovery Procedure:**
1. **Restore from backup**
2. **Verify Azure AD configuration**
3. **Test authentication**
4. **Validate API endpoints**
5. **Monitor logs for errors**

**Recovery Script:**
```bash
#!/bin/bash
# recovery.sh
BACKUP_DIR="/backup/excel-middleware"
LATEST_CONFIG=$(ls -t $BACKUP_DIR/config_*.tar.gz | head -1)
LATEST_LOGS=$(ls -t $BACKUP_DIR/logs_*.tar.gz | head -1)

# Stop service
systemctl stop excel-gpt-middleware

# Restore configuration
tar -xzf $LATEST_CONFIG -C /

# Restore logs
tar -xzf $LATEST_LOGS -C /

# Start service
systemctl start excel-gpt-middleware

# Verify service
sleep 10
curl -f http://localhost:3000/health || exit 1
echo "Recovery completed successfully"
```

## Maintenance

### Regular Maintenance Tasks

**Daily:**
- Check service health
- Monitor error logs
- Verify disk space

**Weekly:**
- Review security logs
- Update dependencies (if needed)
- Performance analysis

**Monthly:**
- Rotate logs
- Update system packages
- Security audit
- Backup verification

### Update Procedure

```bash
#!/bin/bash
# update.sh
set -e

echo "Starting update procedure..."

# Backup current version
cp -r /opt/excel-gpt-middleware /opt/excel-gpt-middleware.backup

# Pull latest code
cd /opt/excel-gpt-middleware
git pull origin main

# Install dependencies
npm ci --only=production

# Run tests (if available)
npm test

# Restart service
systemctl restart excel-gpt-middleware

# Verify service
sleep 10
curl -f http://localhost:3000/health

echo "Update completed successfully"
```

### Log Rotation

```bash
# /etc/logrotate.d/excel-gpt-middleware
/var/log/excel-gpt-middleware/*.log {
    daily
    rotate 30
    compress
    delaycompress
    missingok
    notifempty
    create 640 excel-middleware adm
    postrotate
        systemctl reload excel-gpt-middleware
    endscript
}
```

## Troubleshooting

### Common Issues

**1. Service Won't Start**
```bash
# Check logs
journalctl -u excel-gpt-middleware -f

# Check configuration
node -c /opt/excel-gpt-middleware/src/server.js

# Check permissions
ls -la /opt/excel-gpt-middleware
```

**2. Authentication Failures**
```bash
# Test Azure AD connectivity
curl -X POST https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id={client-id}&client_secret={client-secret}&scope=https://graph.microsoft.com/.default&grant_type=client_credentials"
```

**3. High Memory Usage**
```bash
# Monitor memory usage
ps aux | grep node
top -p $(pgrep node)

# Check for memory leaks
node --inspect /opt/excel-gpt-middleware/src/server.js
```

**4. Performance Issues**
```bash
# Check system resources
htop
iotop
netstat -tulpn

# Analyze logs
grep "duration" /var/log/excel-gpt-middleware/combined-*.log | sort -k4 -nr | head -20
```

### Emergency Procedures

**Service Recovery:**
```bash
# Quick restart
systemctl restart excel-gpt-middleware

# Force kill and restart
pkill -f "node.*server.js"
systemctl start excel-gpt-middleware

# Rollback to previous version
systemctl stop excel-gpt-middleware
rm -rf /opt/excel-gpt-middleware
mv /opt/excel-gpt-middleware.backup /opt/excel-gpt-middleware
systemctl start excel-gpt-middleware
```

**Database/Log Cleanup:**
```bash
# Clean old logs
find /var/log/excel-gpt-middleware -name "*.log" -mtime +7 -delete

# Clear cache (if applicable)
rm -rf /tmp/excel-middleware-cache/*

# Free up disk space
df -h
du -sh /var/log/excel-gpt-middleware/*
```

## Performance Optimization

### Application Tuning

```javascript
// Performance optimizations in server.js
const cluster = require('cluster');
const numCPUs = require('os').cpus().length;

if (cluster.isMaster && process.env.NODE_ENV === 'production') {
    // Fork workers
    for (let i = 0; i < numCPUs; i++) {
        cluster.fork();
    }
    
    cluster.on('exit', (worker, code, signal) => {
        console.log(`Worker ${worker.process.pid} died`);
        cluster.fork();
    });
} else {
    // Worker process
    const server = new Server();
    server.start();
}
```

### System Optimization

```bash
# Increase file descriptor limits
echo "excel-middleware soft nofile 65536" >> /etc/security/limits.conf
echo "excel-middleware hard nofile 65536" >> /etc/security/limits.conf

# Optimize TCP settings
echo "net.core.somaxconn = 65535" >> /etc/sysctl.conf
echo "net.ipv4.tcp_max_syn_backlog = 65535" >> /etc/sysctl.conf
sysctl -p
```

This deployment guide provides comprehensive instructions for deploying the Excel GPT Middleware in production environments with proper security, monitoring, and maintenance procedures.
