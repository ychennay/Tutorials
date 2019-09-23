# Monitoring

## CloudTrail

An **event** in the context of CloudTrail is an action (`ListObjects`) a principal (user `ychen`) performs on an AWS resource (S3 bucket). 

A **trail** is a setup that delivers event logs to a user-specified S3 bucket.

By default, CloudTrail logs **90 days of management events**, with separate vent histories for each region. GLobal events like Route 53 or IAM are included in each region's event history. If you want more than 90 days of event history or want to customize event types, you should create a **trail** (up to 5 trails per region).

### Management Events

Management events are grouped into two types:
* **read-only** events (`DescribeInstances` API call on EC2, any action that reads a resource but cannot modify it)
* **write-only** events (API operations such as logging into management console)

### Data Events

### Log File Integrity
