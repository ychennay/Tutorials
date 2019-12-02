# Authentication and Authorization

IAM is a globally-scoped resource, so you do not need to individually configure IAM policies per region or availability zone. Resource limit of 5,000 IAM users.

## Policies

Policies have three components:
* **Effect**: either `Allow` or `Deny`
* **Action**: a verb
* **Resource** a noun representing a specific AWS resource

If two policies are attached that conflict with each other, the `deny` rule will always overrule the explicit `allow`.

Root users have certain privileges even `AdministratorAccess` policy holders don't, such as:
* ability to create or delete account-wide budgets
* enable MFA Delete on an S3 bucket

## Trusted Entities

4 types of trusted entities:
* AWS service
* another AWS account (identified by its account ID)
* web identity authenticated by logging in w/ Amazon, Cognito, Facebook, or Google
* SAML 2.0 federation with SAML provider

### Workflow

1. User clicks log-in button:
![Login](public/confirm.png)
2. Provider lists requested scopes and permissions (in this case, my profile):
![Confirm](public/login.png)
3. After my credentials are authenticated, a temporary access token is provided:
![Token](public/token.png)
In this case, the token expires in 1 hour, with token type `Bearer` (the standard token type in OAuth 2.0 flows). It also specifies the granted scope as `profile`.

4. Request is sent to the STS endpoint to retrieve temporary AWS credentials:
```
GET / HTTP/1.1
Host: sts.amazonaws.com
Content-Type: application/json; charset=utf-8
URL: https://sts.amazonaws.com/?ProviderId=www.amazon.com&DurationSeconds=900&Action=AssumeRoleWithWebIdentity&Version=2011-06-15&RoleSessionName=web-identity-federation&RoleArn=arn:aws:iam::877950674958:role/WebIdFed_Amazon&WebIdentityToken=Atza|IwEBIEQvRqvge1hx...[OMITTED]...8KU5ltVBQqPSbyNsig
```
The response is an XML object:

```xml
<AssumeRoleWithWebIdentityResponse
  xmlns="https://sts.amazonaws.com/doc/2011-06-15/">
  <AssumeRoleWithWebIdentityResult>
    <Audience>amzn1.application.5d4e68885bb145949aa46756776c27ab</Audience>
    <AssumedRoleUser>
      <AssumedRoleId>AROAJU4SA2VW5SZRF2YMG:web-identity-federation</AssumedRoleId>
      <Arn>arn:aws:sts::ACCOUNT_ID:assumed-role/WebIdFed_Amazon/web-identity-federation</Arn>
    </AssumedRoleUser>
    <Provider>www.amazon.com</Provider>
    <Credentials>
      <AccessKeyId>ACCESS_KEY_ID</AccessKeyId>
      <SecretAccessKey>KEY_TO_SIGN_REQUESTS</SecretAccessKey>
      <SessionToken>FQoGZXIvYX</SessionToken>
      <Expiration>2019-09-21T19:06:04Z</Expiration>
    </Credentials>
    <SubjectFromWebIdentityToken>amzn1.account.AGNE4F4NJ4HUNTAFC4CPD6UKLFUA</SubjectFromWebIdentityToken>
  </AssumeRoleWithWebIdentityResult>
  <ResponseMetadata>
    <RequestId>c3379adb-dca0-11e9-94db-0f5ea0054eb1</RequestId>
  </ResponseMetadata>
</AssumeRoleWithWebIdentityResponse>
```
5. Use the session token to perform resource accesses to AWS resources (in this case a `ListBucket` operation on S3):

```

GET https://web-identity-federation-playground.s3.amazonaws.com/?prefix=user_fun%2F&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=ACCESS_KEY_ID%2F20190921%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-SignedHeaders=host&X-Amz-Date=20190921T185124Z&X-Amz-Security-Token=FQoGZXIvYX&X-Amz-Expires=900&X-Amz-Signature=0ab7e992d28fa6b5f82777aa70bd1795867c605ddbb736bc68b00f7f9b6ddab4
```

## Cognito

* Supports MFA
* User pools - directory profile for all users
* User federation through third-party identity provider
* Signed users receive authentication tokens, which are exchanged for AWS access via **identity pools**

User pools handle authentication.

Identity Pools (authorization):
* assign temporary AWS credentials via AWS STS
* can support anonymous guest users

## Example Use Cases

Customer runs a mobile pet-minder application that needs to be rearchitected to run on AWS. The backend receives pictures every 10 seconds from source cameras and updates the end user's mobile app with the latest picture of their pet. Pictures are not stored; they are either delivered or deleted when the next picture arrives.

The customer wants to make sure of three things: 
1. Only their customers can access the pictures
2. The service can cope with sudden large peaks in load
3. The service is highly available.

**Proposal**:
- Use Amazon Cognito for mobile user sign-in (fulfills requirement 1).
- Use an Auto Scaling group of EC2 instances across multiple Availability Zones to serve the images (fulfills requirement 3)
- Amazon ElastiCache with a Redis replication group to cache, and Multi-AZ RDS to store user information (fulfills requirement 2) and handle the sudden large spikes in load

## CLI Examples

To check last time an access key was used:

`aws iam get-access-key-last-used --access-key-id`

