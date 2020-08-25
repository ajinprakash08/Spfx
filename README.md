yo @microsoft/sharepoint

For deployment
gulp bundle --ship
gulp package-solution --ship

## extensiondemo

This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
# Spfx

##to debug spfx extesion
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"id of ur extension":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}

