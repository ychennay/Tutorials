## Encryption

KMS is global service but keys are regional. AWS KMS assures keys have 99.999999999% durability. 

The **master key** created by AWS KMS on your behalf cannot be exported.

### Customer Master Keys (CMK)

Used often for **envelope encryption**, where you generate, encrypt, and decrpyt the data keys you use to encrypt your data.

CMKs can decrypt/encrypt up to **4Kb of data**.

Types:
* **Customer managed CMKs**: user created, managed, used, rotated
* **AWS managed CMKs**: managed by AWS services, unique to AWS account and region, and only the service that creates the AWS managed CMK can use it
* **default master key**: specific to each service, managed by AWS KMS
* 