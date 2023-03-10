<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>audio.js</title>
    <script src="./audiojs/audio.min.js"></script>
    <link rel="stylesheet" href="./includes/index.css" media="screen">
    <script>
      audiojs.events.ready(function() {
        audiojs.createAll();
      });
    </script>
  </head>
  <body>
    <header>
      <h1>audio.js</h1>
    </header>

    <audio src="http://kolber.github.io/audiojs/demos/mp3/juicy.mp3" preload="auto"></audio>

    <h2>audio.js is a drop-in javascript library that allows HTML5’s &lt;audio&gt; tag to be used anywhere.</h2>

    <p>It uses native &lt;audio&gt; where available and an invisible flash player to emulate &lt;audio&gt; for other browsers. It provides a consistent html player UI to all browsers which can be styled used standard css.</audio>

    <div class="download">
      <a href="http://kolber.github.com/audiojs/audiojs.zip" class="button"><em>Download</em> audio.js</a>
    </div>

    <h3>Installation</h3>
    <ol>
      <li>
        <p>Put <code>audio.js</code>, <code>player-graphics.gif</code> & <code>audiojs.swf</code> in the same folder.</p>
      </li>
      <li>
        <p>Include the <code>audio.js</code> file:</p>
        <pre><code>&lt;script src="/audiojs/audio.min.js"&gt;&lt;/script&gt;</code></pre>
      </li>
      <li>
        <p>Initialise audio.js:</p>
        <pre><code>&lt;script&gt;
  audiojs.events.ready(function() {
    var as = audiojs.createAll();
  });
&lt;/script&gt;</code></pre>
      </li>
      <li>
        <p>Then you can use <code>&lt;audio&gt;</code> wherever you like in your HTML:</p>
        <pre><code>&lt;audio src="/mp3/juicy.mp3" preload="auto" /&gt;</code></pre>
      </li>
    </ol>

    <h3>Examples</h3>
    <p>A series of API tests & examples for using and extending audio.js</p>
    <p><em>Example 1</em> <a href="http://kolber.github.com/audiojs/demos/test1.html">Test multiple load types</a></p>
    <p><em>Example 2</em> <a href="http://kolber.github.com/audiojs/demos/test2.html">Custom markup/css</a></p>
    <p><em>Example 3</em> <a href="http://kolber.github.com/audiojs/demos/test3.html">Multiple players, testing <code>preload</code>, <code>loop</code> & <code>autoplay</code> attributes</a></p>
    <p><em>Example 4</em> <a href="http://kolber.github.com/audiojs/demos/test5.html">Customised player</a></p>
    <p><em>Example 5</em> <a href="http://kolber.github.com/audiojs/demos/test6.html">Customised playlist player</a></p>
    <!--<p><em>Example 6</em> <a href="http://kolber.github.com/audiojs/demos/test4.html">.ogg fallback</a></p>-->

    <h3>Browser & format support</h3>
    <p>With Flash as a fallback, it should work pretty much anywhere.<br>
      It has been verified to work across:</p>
    <ul>
      <li>Mobile Safari <em>(iOS 3+)</em></li>
      <li>Android <em>(2.2+, w/Flash)</em></li>
      <li>Safari <em>(4+)</em></li>
      <li>Chrome <em>(7+)</em></li>
      <li>Firefox <em>(3+, w/ Flash)</em></li>
      <li>Opera <em>(10+, w/ Flash)</em></li>
      <li>IE <em>(6, 7, 8, w/ Flash)</em></li>
    </ul>
    <p><strong>ogg</strong></p>
    <p>audio.js focuses on playing mp3s. It doesn’t currently support the ogg format. As mp3 is the current defacto music transfer format, ogg support is lower on our list of priorities.</p>

    <h3>Flash local security</h3>
    <blockquote>
      <p><strong>Note:</strong> For local content running in a browser, calls to the <code>ExternalInterface.addCallback()</code> method work only if the SWF file and the containing web page are in the local-trusted security sandbox.</p>
      <cite><a href="http://help.adobe.com/en_US/FlashPlatform/reference/actionscript/3/flash/external/ExternalInterface.html#addCallback()">http://help.adobe.com/en_US/FlashPlatform/reference/actionscript/3/flash/external/ExternalInterface.html#addCallback()</a></cite>
    </blockquote>
    <p>This means that unless you have gone through the rigmarole of setting up your <a href="http://kb2.adobe.com/cps/093/4c093f20.html#main_blocked">flash player security settings for local files</a>, <code>ExternalInterface</code> calls will only work when the page is loaded from a ‘domain’. <code>http://localhost</code> counts, but any <code>file://</code> requests don’t.</p>

    <h3>Source code</h3>
    <p>All efforts have been made to keep the source as clean and readable as possible. Until we release more detailed documentation, the annotated source is the best reference for usage.</p>
    <p><a href="http://kolber.github.com/audiojs/docs/">Annotated source</a> / <a href="http://github.com/kolber/audiojs">Source on Github</a></p>

    <h3>License</h3>
    <p>audio.js is released under an <a href="http://www.opensource.org/licenses/mit-license.php">MIT License</a>, so do with it what you will.</p>
    <footer>
      <p>All example audio files are from <a href="http://waitwhatmusic.com/">wait what</a>’s <a href="http://soundcloud.com/wait-what/sets/the-notorious-xx">notorious xx album</a> and used with permission.</p>
      <p>This site is ©copyright <a href="http://aestheticallyloyal.com">Anthony Kolber</a>, 2010.</p>
      <p class="ab-c"><em>&real;</em> Another <a href="http://ab-c.com.au">ab+c</a> joint</p>
    </footer>
  </body>
</html>
