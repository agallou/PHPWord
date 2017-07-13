<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpWord\Element\Image as ImageElement;
use PhpOffice\PhpWord\Writer\Word2007\Style\Image as ImageStyleWriter;

/**
 * Image element writer
 *
 * @since 0.10.0
 */
class Image extends AbstractElement
{
    /**
     * Write element.
     *
     * @return void
     */
    public function write()
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof ImageElement) {
            return;
        }

        $this->writeImage($xmlWriter, $element);
    }

    /**
     * Write image element.
     *
     * @return void
     */
    private function writeImage(XMLWriter $xmlWriter, ImageElement $element)
    {
        $rId = $element->getRelationId() + ($element->isInSection() ? 6 : 0);
        $style = $element->getStyle();

        $styleWriter = new ImageStyleWriter($xmlWriter, $style);

        $cx = \PhpOffice\PhpWord\Shared\Converter::pixelToEmu($style->getWidth());
        $cy = \PhpOffice\PhpWord\Shared\Converter::pixelToEmu($style->getHeight());

        $xmlWriter->startElement('w:p');
        $styleWriter->writeAlignment();
        $xmlWriter->startElement('w:r');

        $xmlWriter->startElement('w:drawing');

        $xmlWriter->startElement('wp:inline');
        $xmlWriter->writeAttribute('distT', 0);
        $xmlWriter->writeAttribute('distB', 0);
        $xmlWriter->writeAttribute('distL', 114300);
        $xmlWriter->writeAttribute('distR', 114300);

        $xmlWriter->startElement('wp:extent');
        $xmlWriter->writeAttribute('cx', $cx);
        $xmlWriter->writeAttribute('cy', $cy);
        $xmlWriter->endElement(); //wp:extent

        $xmlWriter->startElement('wp:docPr');
        $xmlWriter->writeAttribute('id', $rId);
        $xmlWriter->writeAttribute('name', 'name');
        $xmlWriter->writeAttribute('descr', 'aa');
        $xmlWriter->endElement(); // wp:docPr

        $xmlWriter->startElement('wp:cNvGraphicFramePr');
        $xmlWriter->startElement('a:graphicFrameLocks');
        $xmlWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xmlWriter->writeAttribute('noChangeAspect', 1);
        $xmlWriter->endElement(); //a:graphicFrameLocks
        $xmlWriter->endElement(); // wp:cNvGraphicFramePr

        $xmlWriter->startElement('a:graphic');
        $xmlWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $xmlWriter->startElement('a:graphicData');
        $xmlWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
        $xmlWriter->startElement('pic:pic');
        $xmlWriter->writeAttribute('xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture');


        $xmlWriter->startElement('pic:nvPicPr');
        $xmlWriter->startElement('pic:cNvPr');
        $xmlWriter->writeAttribute('id', 3);
        $xmlWriter->writeAttribute('desc', 'aa');
        $xmlWriter->writeAttribute('name', 'name');
        $xmlWriter->endElement(); //pic:cNvPr
        $xmlWriter->startElement('pic:cNvPicPr');
        $xmlWriter->startElement('a:picLocks');
        $xmlWriter->writeAttribute('noChangeAspect', '1');
        $xmlWriter->endElement(); //a:picLocks
        $xmlWriter->endElement(); //pic:cNvPicPr
        $xmlWriter->endElement(); //pic:nvPicPr


        $xmlWriter->startElement('pic:blipFill');

        $xmlWriter->startElement('a:blip');
        $xmlWriter->writeAttribute('r:embed', 'rId' .$rId);
        $xmlWriter->endElement(); //a:blip

        $xmlWriter->startElement('a:stretch');
        $xmlWriter->startElement('a:fillRect');
        $xmlWriter->endElement(); //a:fillRect
        $xmlWriter->endElement(); //a:stretch

        $xmlWriter->endElement(); //pic:blipFill

        $xmlWriter->startElement('pic:spPr');
        $xmlWriter->startElement('a:xfrm');

        $xmlWriter->startElement('a:off');
        $xmlWriter->writeAttribute('x', 0);
        $xmlWriter->writeAttribute('y', 0);
        $xmlWriter->endElement(); //a:off

        $xmlWriter->startElement('a:ext');
        $xmlWriter->writeAttribute('cx', $cx);
        $xmlWriter->writeAttribute('cy', $cy);
        $xmlWriter->endElement(); //a:ext

        $xmlWriter->endElement(); //a:xfrm

        $xmlWriter->startElement('a:prstGeom');
        $xmlWriter->writeAttribute('prst', 'rect');
        $xmlWriter->startElement('a:avLst');
        $xmlWriter->endElement(); //a:avLst
        $xmlWriter->endElement(); //a:prstGeom

        $xmlWriter->endElement(); //pic:spPr

        $xmlWriter->endElement(); //pic:pic
        $xmlWriter->endElement(); //a:graphicData
        $xmlWriter->endElement(); //a:graphic
        $xmlWriter->endElement(); //wp:inline
        $xmlWriter->endElement(); //w:drawing

        $xmlWriter->endElement(); //w:r
        $xmlWriter->endElement(); //w:p
    }
}

