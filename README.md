[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]

# Automated Outlook Signature Scripts
This project currently contains one script which can be utilised to automate the process of setting users Outlook signatures.
This script has been tested with Outlook 2010, 2016, 2019, 2021.

Currently I recommend using the script in Group Policy as a logon script. If you're unfamiliar with this process, you can follow the detailed instructions provided in this article - [Configuring Logon PowerShell Scripts with Group Policy - 4Sysops](https://4sysops.com/archives/configuring-logon-powershell-scripts-with-group-policy/) - I ahve also created a short video guide: [YouTube Video Guide](https://www.youtube.com/watch?v=rt9y02iBoPE)

If configured this way, during the user's logon process, the script runs in the background, retrieves the necessary user details, generates a new signature file, and replaces the existing one. Additionally, the script sets registry keys to configure the newly created signature as the user's default Outlook signature. This ensures that if any details such as job title change, the signature will be automatically updated during the next logon.

I posted this script on EduGeek forum many moons ago so any additional conversation can be had over there. [EduGeek Post](http://www.edugeek.net/forums/scripts/205976-outlook-email-signature-automation-ad-attributes.html#post1760284)

### Launch Parameters



[contributors-shield]: https://img.shields.io/github/contributors/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[contributors-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[forks-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/network/members
[stars-shield]: https://img.shields.io/github/stars/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[stars-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/stargazers
[issues-shield]: https://img.shields.io/github/issues/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[issues-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/issues
[license-shield]: https://img.shields.io/github/license/CaptainQwerty/AutomatedOutlookSignature.svg?style=for-the-badge
[license-url]: https://github.com/CaptainQwerty/AutomatedOutlookSignature/blob/master/LICENSE.txt