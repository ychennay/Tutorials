## Simple Notification Service

A **push notification, pub-sub** service. 

### Topics

- must have a unique name within AWS account
- at most 256 alphanumeric characters

### Delivery Formats

- **Email**: traditional user emails
- **SMS**: registered phone number
- **Email (JSON)**: applications that process emails
- **HTTP/HTTPs**: URL specified, notification is sent as `POST`.
- **SQS**: enqueue messages into SQS queue.
- **Lambda**: trigger serverless function