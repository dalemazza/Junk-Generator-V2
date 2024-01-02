<!-- Improved compatibility of back to top link: See: https://github.com/othneildrew/Best-README-Template/pull/73 -->
<a name="readme-top"></a>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->


<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/KhaosShield/Junk-Generator-V2">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

<h3 align="center">Junk Generator</h3>

  <p align="center">
    A Python script that allows the user to fill Microsoft Office files s with 'Junk' while not impacting any existing Macros.
    <br />
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

[![Product Name Screen Shot][product-screenshot]](https://example.com)

Junk Generator will simply fill Microsoft Documents with junk text, to alter file size to something more believable. The aim is to simply trigger any macros contained in the document. The script currently works on the following documents:
* []() .docx
* []() .docm
* []() .pptx
* []() .pptm
* []() .xlsm

### Word
For .docx files, it uses python-docx to add paragraphs of junk text.
For .docm files, it uses COM automation with Microsoft Word to add text. This approach ensures that macros within the .docm files are preserved.

### Powerpoint
The script checks if the PowerPoint file (.pptx or .pptm) is empty and adds a title slide if necessary.
It then adds junk text to the notes section of each slide.

### Excel Workbooks
For .xlsm files, the script uses openpyxl to add junk text to the first 10 rows of the active worksheet.




<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- GETTING STARTED -->
## Getting Started

To get start simply run the script and provide the $PATH to the Document you wish to add junk too.

### Prerequisites

Dependencies
* 
  ```sh
  pip install python-pptx openpyxl python-docx pywin32
  ```

### Installation

1. Clone the repo
   ```sh
   git clone https://github.com/KhaosShield/Junk-Generator-V2.git
   ```
2. Install Dependencies
   ```sh
   pip install python-pptx openpyxl python-docx pywin32
   ```
3. Run the script
   ```sh
   ./JunkGen.py
   ```

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- USAGE EXAMPLES -->
## Usage

Simply run JunkGen.py and provide the full path to the file you wish to add junk too. Easy.



<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## Roadmap

- [ ] Make better
- [ ] 


See the [open issues](https://github.com/KhaosShield/Junk-Generator-V2/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Your Name - [@KhaosShield](https://twitter.com/@KhaosShield) - khaosshield@protonmail.com

Project Link: [https://github.com/KhaosShield/Junk-Generator-V2](https://github.com/KhaosShield/Junk-Generator-V2)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ACKNOWLEDGMENTS -->
## Acknowledgments

* []()
* []()
* []()

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/KhaosShield/Junk-Generator-V2.svg?style=for-the-badge
[contributors-url]: https://github.com/KhaosShield/Junk-Generator-V2/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/KhaosShield/Junk-Generator-V2.svg?style=for-the-badge
[forks-url]: https://github.com/KhaosShield/Junk-Generator-V2/network/members
[stars-shield]: https://img.shields.io/github/stars/KhaosShield/Junk-Generator-V2.svg?style=for-the-badge
[stars-url]: https://github.com/KhaosShield/Junk-Generator-V2/stargazers
[issues-shield]: https://img.shields.io/github/issues/KhaosShield/Junk-Generator-V2.svg?style=for-the-badge
[issues-url]: https://github.com/KhaosShield/Junk-Generator-V2/issues
[license-shield]: https://img.shields.io/github/license/KhaosShield/Junk-Generator-V2.svg?style=for-the-badge
[license-url]: https://github.com/KhaosShield/Junk-Generator-V2/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/linkedin_username
[product-screenshot]: images/screenshot.png
[Next.js]: https://img.shields.io/badge/next.js-000000?style=for-the-badge&logo=nextdotjs&logoColor=white
[Next-url]: https://nextjs.org/
[React.js]: https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB
[React-url]: https://reactjs.org/
[Vue.js]: https://img.shields.io/badge/Vue.js-35495E?style=for-the-badge&logo=vuedotjs&logoColor=4FC08D
[Vue-url]: https://vuejs.org/
[Angular.io]: https://img.shields.io/badge/Angular-DD0031?style=for-the-badge&logo=angular&logoColor=white
[Angular-url]: https://angular.io/
[Svelte.dev]: https://img.shields.io/badge/Svelte-4A4A55?style=for-the-badge&logo=svelte&logoColor=FF3E00
[Svelte-url]: https://svelte.dev/
[Laravel.com]: https://img.shields.io/badge/Laravel-FF2D20?style=for-the-badge&logo=laravel&logoColor=white
[Laravel-url]: https://laravel.com
[Bootstrap.com]: https://img.shields.io/badge/Bootstrap-563D7C?style=for-the-badge&logo=bootstrap&logoColor=white
[Bootstrap-url]: https://getbootstrap.com
[JQuery.com]: https://img.shields.io/badge/jQuery-0769AD?style=for-the-badge&logo=jquery&logoColor=white
[JQuery-url]: https://jquery.com 