# Configure-ClientSizeLimits.ps1

Configure client-specific message size limits for Exchange Web Services, Outlook WebApp
or ActiveSync workloads. Can run locally or remotely, from the Exchange Management Shell.
Specified limits are in 1KB units. 
More information at https://technet.microsoft.com/en-us/library/hh529949%28v=exchg.150%29.aspx

## Prerequisites

Script requires Exchange Management Shell, and works against Exchange 2013/2016.
	
## Usage

```
Configure-ClientSizeLimits.ps1 -OWA 25MB -EWS 15MB -EAS 25MB
```
Configure client size limit of 25MB for OWA, 15MB for EWS and 25MB for ActiveSync.

## Contributing

N/A

## Versioning

Initial version published on GitHub is 1.2. Changelog is contained in the script.

## Authors

* Michel de Rooij [initial work] https://github.com/michelderooij

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

## Acknowledgments

N/A
 