# Monitoring

## Log types

### VPC Flow Logs

### CloudWatch Logs
- aggregator for any type of other log (including on-premises)
- often collected via CloudWatch agents on EC2 instances

### AWS Config
- Inventory of AWS resources and record of resource changes
- Troubleshoot outages, conduct security attack analysis

## CloudTrail

An **event** in the context of CloudTrail is an action (`ListObjects`) a principal (user `ychen`) performs on an AWS resource (S3 bucket). 

A **trail** is a setup that delivers event logs to a user-specified S3 bucket. A trail can be configured for either a specific region, or all regions.

By default, CloudTrail logs **90 days of management events**, with separate vent histories for each region. Global events like Route 53 or IAM are included in each region's event history. If you want more than 90 days of event history or want to customize event types, you should create a **trail** (up to 5 trails per region). A trail that is global will count as one trail per region.

Creating a trail via the management console defaults to **all regions**. On the API or CLI, the default is a **single region.**

### Management Events

Management events are grouped into two types:
* **read-only** events (`DescribeInstances` API call on EC2, any action that reads a resource but cannot modify it)
* **write-only** events (API operations such as logging into management console)

### Data Events

### Log File Integrity

### Sample Use Cases

#### Monitoring EC2 State Changes

Customer's CloudWatch Logs configuration receives logs and data from on-premises monitoring systems and agents installed in operating systems. A new team wants to use CloudWatch to also monitor Amazon EC2 instance performance and state changes of EC2 instances, such as instance creation, instance power-off, and instance termination. This solution should also be able to notify the team of any state changes for troubleshooting. What should you do to monitor EC2 instance state changes and performance?

Proposed:
- Configure a CloudWatch Events rule for EC2 instance state changes. Have the Events rule route to an Amazon SNS topic that notifies your client.

Considerations:
- EC2 Instance state changes do not trigger CloudWatch alarms
- Do not need an event rule for all API calls, only state changes! Otherwise you'll have to do a ton of filtering in CloudWatch, or receive a ton of emails
- Amazon Inspector would not capture state changes in EC2 instances

#### On-Premise Monitoring Integration
Your customer’s **existing on-premises monitoring system** currently receives performance notifications from local management systems and data from agents installed in on-premises operating systems. The customer wants to use this on-premises monitoring system to **monitor Amazon EC2 instance performance and state changes of EC2 instances**.

What should you do to **monitor EC2 instance state changes and performance**?

Proposed:
1. Configure a **CloudWatch Events rule** for Amazon EC2 instance state changes. 
2. Have the Events rule **trigger an AWS Lambda function** that notifies your monitoring system. 
3. **Pull CloudWatch metric data** into your monitoring system.

Considerations:
- CloudTrail is for API calls, and is not a relevant service for this problem (we are worried about performance monitoring and state transitions, not user API calls)

#### Monitoring via Kibana

Client uses Kibana to create dashboards and visualizations of their on-premises environment to benchmark the health of their systems. They use this information to reactively respond to events when needed. They also want to use Kibana for all of their AWS VPC and Amazon EC2 log data, but require an automated means of responding to events in the cloud. Which solution would meet the requirements set by your client?

**Proposal:** Use CloudWatch Logs to gather logs and Lambda for automated responses. Ingest log data into Amazon ES (ElasticSearch) and use its built-in support for Kibana.

Consideration:
- **Kinesis Data Analytics** and **AWS Athena** would not integrate w/ Kibana.
- With **Kinesis Streams**, you build applications using the Kinesis Producer Library put the data into a stream and then process it with an application that uses the Kinesis Client Library and with Kinesis Connector Library send the processed data to S3, Redshift, DynamoDB etc.
- With **Kinesis Firehose** it’s a bit simpler: create delivery stream and send the data to S3, Redshift or ElasticSearch (using the Kinesis Agent or API) directly and storing it in those services.
- Use Kinesis Streams if you want to do some custom processing with streaming data. With Kinesis Firehose you are simply ingesting it into S3, Redshift or ElasticSearch.