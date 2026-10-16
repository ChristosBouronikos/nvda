# A part of NonVisual Desktop Access (NVDA)
# Copyright (C) 2026 NV Access Limited, Christos Bouronikos
# This file may be used under the terms of the GNU General Public License, version 2 or later, as modified by the NVDA license.
# For full terms and any additional permissions, see the NVDA license file: https://github.com/nvaccess/nvda/blob/master/copying.txt

"""Tests for PowerPoint slide title and content reporting."""

import unittest
from collections.abc import Iterable
from types import SimpleNamespace
from typing import Any
from unittest.mock import Mock, PropertyMock, patch

import comtypes
import textInfos
from appModules import powerpnt


class _Shapes(list[SimpleNamespace]):
	def __init__(self, shapes: Iterable[SimpleNamespace] = (), title: SimpleNamespace | None = None) -> None:
		super().__init__(shapes)
		self._title = title

	@property
	def title(self) -> SimpleNamespace:
		if self._title is None:
			raise comtypes.COMError(-1, "No title", None)
		return self._title


def _textShape(shapeID: int, text: str) -> SimpleNamespace:
	paragraphs = [
		SimpleNamespace(text=paragraph, paragraphFormat=SimpleNamespace(bullet=SimpleNamespace(type=0)))
		for paragraph in text.split("\r")
	]
	return SimpleNamespace(
		ID=shapeID,
		type=powerpnt.msoTextBox,
		hasTextFrame=True,
		alternativeText="",
		title="",
		textFrame=SimpleNamespace(textRange=SimpleNamespace(text=text, paragraphs=lambda: paragraphs)),
	)


def _object[T](cls: type[T], **attrs: Any) -> T:
	obj = object.__new__(cls)
	obj._propertyCache = {}
	obj.__dict__.update(attrs)
	return obj


class TestSlideTitles(unittest.TestCase):
	def setUp(self) -> None:
		self.titleShape = _textShape(1, "Παράδειγμα")
		self.bodyShape = _textShape(2, "Body text")
		self.ppSlide = SimpleNamespace(
			slideNumber=1,
			slideID=101,
			shapes=_Shapes([self.titleShape, self.bodyShape], title=self.titleShape),
		)
		self.document = _object(
			powerpnt.DocumentWindow,
			ppActivePaneViewType=powerpnt.ppViewSlide,
		)
		self.slide = _object(powerpnt.Slide, ppObject=self.ppSlide, documentWindow=self.document)
		self.document.currentSlide = self.slide
		self.window = _object(powerpnt.SlideShowWindow, currentSlide=self.slide, notesMode=False)

	def test_slideNameLabelsTitleWithoutParentheses(self) -> None:
		self.assertEqual(self.slide.name, "Slide 1 - Slide title: Παράδειγμα")

	def test_documentContextDoesNotRepeatTitle(self) -> None:
		self.assertNotIn("Παράδειγμα", self.document.name)
		self.assertIn("Slide 1", self.document.name)

	def test_slideShowTitleOccursOnceAndStaysReviewable(self) -> None:
		self.assertNotIn("Παράδειγμα", self.window.name)
		self.assertEqual(
			self.window.basicText,
			"Slide title: Παράδειγμα\nSlide contents\nBody text",
		)

	def test_titleIsReportedFirstEvenWhenNotFirstShape(self) -> None:
		self.ppSlide.shapes.reverse()
		self.assertTrue(self.window.basicText.startswith("Slide title: Παράδειγμα\n"))

	def test_sameTextInAnotherShapeIsNotRemoved(self) -> None:
		self.bodyShape.textFrame.textRange.paragraphs = self.titleShape.textFrame.textRange.paragraphs
		self.assertEqual(self.window.basicText.count("Παράδειγμα"), 2)

	def test_titleOnlySlideDoesNotAnnounceEmptyContents(self) -> None:
		self.ppSlide.shapes = _Shapes([self.titleShape], title=self.titleShape)
		self.assertEqual(self.window.basicText, "Slide title: Παράδειγμα")

	def test_untitledSlideStillReportsContents(self) -> None:
		self.ppSlide.shapes = _Shapes([self.bodyShape])
		self.assertEqual(self.window.basicText, "Slide contents\nBody text")
		self.assertEqual(self.slide.name, "Slide 1")

	def test_emptySlideKeepsEmptySlideMessage(self) -> None:
		self.ppSlide.shapes = _Shapes()
		self.assertEqual(self.window.basicText, "Empty slide")

	def test_blankTitleDoesNotHideContent(self) -> None:
		self.titleShape.textFrame.textRange.text = "\r "
		self.titleShape.textFrame.textRange.paragraphs = list
		self.assertEqual(self.window.basicText, "Slide contents\nBody text")

	def test_notesKeepSlideTitleAndDoNotAddContentLabels(self) -> None:
		self.window.notesMode = True
		notesShape = _textShape(3, "Speaker notes")
		notesShape.placeholderFormat = SimpleNamespace(type=powerpnt.ppPlaceholderBody)
		self.ppSlide.notesPage = SimpleNamespace(shapes=_Shapes([notesShape]))
		self.assertIn("Slide title: Παράδειγμα", self.window.name)
		self.assertEqual(self.window.basicText, "Speaker notes")

	def test_mathFieldOffsetsIncludeSectionLabels(self) -> None:
		field = textInfos.ControlField(role="math")
		with patch.object(
			powerpnt.SlideShowWindow,
			"_getShapeText",
			return_value=iter([powerpnt._SlideTextChunk(field, self.bodyShape)]),
		):
			self.assertEqual(self.window.basicText, "Slide title: Παράδειγμα\nSlide contents\n ")
		self.assertEqual(
			self.window.basicTextFields,
			[(len("Slide title: Παράδειγμα\nSlide contents\n"), field)],
		)

	def test_greekSectionLabels(self) -> None:
		translations = {
			"Slide title: {title}": "Τίτλος διαφάνειας: {title}",
			"Slide contents": "Περιεχόμενα διαφάνειας",
		}
		with patch.object(powerpnt, "_", side_effect=lambda text: translations.get(text, text), create=True):
			self.assertEqual(
				self.window.basicText,
				"Τίτλος διαφάνειας: Παράδειγμα\nΠεριεχόμενα διαφάνειας\nBody text",
			)

	def test_editingTitleShapeUsesSlideTitleLabel(self) -> None:
		shape = _object(powerpnt.Shape, ppObject=self.titleShape, documentWindow=self.document)
		self.assertEqual(shape.name, "Slide title")
		self.assertEqual(shape.value, "Παράδειγμα")

	def test_editingAnnouncesContentsOnSectionTransition(self) -> None:
		title = _object(powerpnt.Shape, ppObject=self.titleShape, documentWindow=self.document)
		body = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		with (
			patch.object(powerpnt.Window, "reportFocus"),
			patch.object(powerpnt.speech, "speakMessage") as speakMessage,
		):
			self.slide.reportFocus()
			title.reportFocus()
			speakMessage.assert_not_called()
			body.reportFocus()
			body.reportFocus()
			speakMessage.assert_called_once_with("Slide contents")
			title.reportFocus()
			body.reportFocus()
			self.assertEqual(speakMessage.call_count, 2)

	def test_editingDirectTextFrameEntryReportsContents(self) -> None:
		parent = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		frame = _object(powerpnt.TextFrame, documentWindow=self.document, parent=parent)
		with (
			patch.object(powerpnt.Window, "reportFocus"),
			patch.object(powerpnt.speech, "speakMessage") as speakMessage,
		):
			frame.reportFocus()
			speakMessage.assert_called_once_with("Slide contents")

	def test_editingTableCellEntryReportsContents(self) -> None:
		table = _object(powerpnt.Table, ppObject=_textShape(3, ""), documentWindow=self.document)
		cell = _object(powerpnt.TableCell, parent=table, table=table, documentWindow=self.document)
		frame = _object(powerpnt.TableCellTextFrame, documentWindow=self.document, parent=cell)
		with (
			patch.object(powerpnt.Window, "reportFocus"),
			patch.object(powerpnt.speech, "speakMessage") as speakMessage,
		):
			frame.reportFocus()
			speakMessage.assert_called_once_with("Slide contents")

	def test_editingNewSlideReportsContentsAgain(self) -> None:
		body = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		with (
			patch.object(powerpnt.Window, "reportFocus"),
			patch.object(powerpnt.speech, "speakMessage") as speakMessage,
		):
			body.reportFocus()
			self.ppSlide.slideID += 1
			body.reportFocus()
			self.assertEqual(speakMessage.call_count, 2)

	def test_notesEditingDoesNotAnnounceSlideContents(self) -> None:
		self.document.ppActivePaneViewType = powerpnt.ppViewNotesPage
		body = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		with (
			patch.object(powerpnt.Window, "reportFocus"),
			patch.object(powerpnt.speech, "speakMessage") as speakMessage,
		):
			body.reportFocus()
			speakMessage.assert_not_called()

	def test_slideChangeStillDetectsTitleEdits(self) -> None:
		self.window._lastSlideChangeID = None
		with (
			patch.object(
				powerpnt.SlideShowWindow,
				"treeInterceptor",
				new_callable=PropertyMock,
			) as treeInterceptor,
			patch.object(powerpnt.SlideShowWindow, "currentSlide", new_callable=PropertyMock) as currentSlide,
			patch.object(powerpnt.SlideShowWindow, "reportFocus") as reportFocus,
		):
			treeInterceptor.return_value = Mock()
			currentSlide.return_value = self.slide
			self.window.handleSlideChange()
			self.window.handleSlideChange()
			reportFocus.assert_called_once()
			self.titleShape.textFrame.textRange.text = "Changed title"
			self.slide.invalidateCache()
			self.window.invalidateCache()
			self.window.handleSlideChange()
			self.assertEqual(reportFocus.call_count, 2)
			self.assertEqual(treeInterceptor.return_value.reportNewSlide.call_count, 2)

	def test_bulletAndParagraphContentIsPreserved(self) -> None:
		paragraph = SimpleNamespace(
			text="First\x0bSecond\r",
			paragraphFormat=SimpleNamespace(bullet=SimpleNamespace(type=powerpnt.ppBulletNumbered, number=2)),
		)
		self.bodyShape.textFrame.textRange.paragraphs = Mock(return_value=[paragraph])
		self.assertTrue(self.window.basicText.endswith("2. First\nSecond"))

	def test_slideTextIsBuiltBeforeAccessingItsFields(self) -> None:
		info = _object(
			powerpnt.SlideShowTreeInterceptorTextInfo,
			obj=SimpleNamespace(rootNVDAObject=self.window),
			_startOffset=0,
			_endOffset=100,
		)
		self.assertEqual(info.getTextWithFields(), ["Slide title: Παράδειγμα\nSlide contents\nBody text"])

	def test_completedSlideShowDoesNotKeepOldFields(self) -> None:
		self.window.currentSlide = None
		self.window.basicTextFields = [(5, textInfos.ControlField(role="math"))]
		_ = self.window.basicText
		self.assertEqual(self.window.basicTextFields, [])

	def test_distinctSlidesWithIdenticalNamesAreReported(self) -> None:
		with (
			patch.object(
				powerpnt.SlideShowWindow,
				"treeInterceptor",
				new_callable=PropertyMock,
			) as interceptor,
			patch.object(powerpnt.SlideShowWindow, "currentSlide", new_callable=PropertyMock) as currentSlide,
			patch.object(powerpnt.SlideShowWindow, "reportFocus") as reportFocus,
		):
			interceptor.return_value = Mock()
			currentSlide.return_value = self.slide
			self.window.handleSlideChange()
			self.ppSlide.slideID += 1
			self.window.handleSlideChange()
			self.assertEqual(reportFocus.call_count, 2)

	def test_slideRefreshCanBeSilentWhenRegainingFocus(self) -> None:
		with (
			patch.object(
				powerpnt.SlideShowWindow,
				"treeInterceptor",
				new_callable=PropertyMock,
			) as interceptor,
			patch.object(powerpnt.SlideShowWindow, "currentSlide", new_callable=PropertyMock) as currentSlide,
			patch.object(powerpnt.SlideShowWindow, "reportFocus") as reportFocus,
		):
			interceptor.return_value = Mock()
			currentSlide.return_value = self.slide
			self.assertTrue(self.window.handleSlideChange(report=False))
			self.assertFalse(self.window.handleSlideChange(report=False))
			reportFocus.assert_not_called()
			interceptor.return_value.reportNewSlide.assert_not_called()

	def test_shapeValueIncludesBulletsWhenTabbingInEditingView(self) -> None:
		self.bodyShape.textFrame.textRange.paragraphs = lambda: [
			SimpleNamespace(
				text="Item\r",
				paragraphFormat=SimpleNamespace(
					bullet=SimpleNamespace(type=powerpnt.ppBulletNumbered, number=3),
				),
			),
		]
		shape = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		self.assertEqual(shape.value, "3. Item")

	def test_smartArtNodesAreReadInModelOrder(self) -> None:
		shape = SimpleNamespace(
			type=powerpnt.msoSmartArt,
			hasTextFrame=False,
			smartArt=SimpleNamespace(
				allNodes=[
					SimpleNamespace(textFrame2=SimpleNamespace(textRange=SimpleNamespace(text=text)))
					for text in ["Start", "", "Review\rFinish"]
				],
			),
		)
		self.assertEqual(
			[chunk.content for chunk in self.window._getShapeText(shape)],
			["Start", "Review\nFinish"],
		)
		obj = _object(powerpnt.Shape, ppObject=shape, documentWindow=self.document)
		self.assertEqual(obj.value, "Start\nReview\nFinish")

	def test_unlabelledMediaStillHasAnAccessibleLabel(self) -> None:
		for mediaType, role in [(powerpnt.ppVideo, "video"), (powerpnt.ppAudio, "audio")]:
			with self.subTest(mediaType=mediaType):
				shape = SimpleNamespace(
					type=powerpnt.msoMedia,
					mediaType=mediaType,
					hasTextFrame=False,
					hasTable=False,
					alternativeText="",
				)
				self.assertEqual([chunk.content for chunk in self.window._getShapeText(shape)], [role])

	def test_decorativeShapesAreSkippedOnlyForSlideShowReading(self) -> None:
		self.bodyShape.decorative = -1
		self.assertEqual(list(self.window._getShapeText(self.bodyShape)), [])
		obj = _object(powerpnt.Shape, ppObject=self.bodyShape, documentWindow=self.document)
		self.assertEqual(obj.value, "Body text")

	def test_symbolGlyphUsesAuthoredAlternativeText(self) -> None:
		self.bodyShape.alternativeText = "cycle symbol"
		self.bodyShape.textFrame.textRange.font = SimpleNamespace(name="Webdings")
		self.assertEqual(
			[chunk.content for chunk in self.window._getShapeText(self.bodyShape)],
			["cycle symbol"],
		)

	def test_notesBodyIsFoundByTypeAndFooterIsNotRead(self) -> None:
		footer = _textShape(4, "Footer")
		footer.placeholderFormat = SimpleNamespace(type=powerpnt.ppPlaceholderFooter)
		notes = _textShape(5, "Actual notes")
		notes.placeholderFormat = SimpleNamespace(type=powerpnt.ppPlaceholderBody)
		self.ppSlide.notesPage = SimpleNamespace(shapes=[footer, self.bodyShape, notes])
		self.window.notesMode = True
		self.assertIs(powerpnt.getNotesTextFrame(self.ppSlide), notes.textFrame)
		self.assertEqual(self.window.basicText, "Actual notes")


class TestPowerPointObjectModelSelection(unittest.TestCase):
	def test_inactiveSlideShowIsUsedForPresenterView(self) -> None:
		show = SimpleNamespace(active=False)
		editor = SimpleNamespace()
		presentation = SimpleNamespace(
			slideShowWindow=show,
			windows=SimpleNamespace(item=lambda index: editor),
		)
		app = SimpleNamespace(_ppApplicationFromROT=SimpleNamespace(activePresentation=presentation))
		self.assertIs(powerpnt.AppModule._getPpObjectModelFromROT(app, slideShow=True), show)
		self.assertIs(powerpnt.AppModule._getPpObjectModelFromROT(app), editor)

	def test_notesSelectionDoesNotDependOnSecondShape(self) -> None:
		notes = _textShape(5, "Notes")
		notes.placeholderFormat = SimpleNamespace(type=powerpnt.ppPlaceholderBody)
		slide = SimpleNamespace(
			notesPage=SimpleNamespace(shapes=[_textShape(1, ""), _textShape(2, ""), notes]),
		)
		document = _object(
			powerpnt.DocumentWindow,
			ppActivePaneViewType=powerpnt.ppViewNotesPage,
			ppDocumentViewType=powerpnt.ppViewNormal,
			ppVersionMajor=16,
			ppSelection=SimpleNamespace(type=powerpnt.ppSelectionText, slideRange={1: slide}),
		)
		with patch.object(powerpnt, "NotesTextFrame") as cls:
			self.assertIs(document.selection, cls.return_value)
		cls.assert_called_once_with(
			windowHandle=document.windowHandle,
			documentWindow=document,
			ppObject=notes.textFrame,
		)


class TestSlideShowFocus(unittest.TestCase):
	def setUp(self) -> None:
		self.window = Mock()
		self.interceptor = _object(powerpnt.SlideShowTreeInterceptor, rootNVDAObject=self.window)

	def test_nativeChildFocusIsNotSuppressed(self) -> None:
		nextHandler = Mock()
		self.interceptor.event_gainFocus(SimpleNamespace(), nextHandler)
		nextHandler.assert_called_once_with()

	def test_rootFocusIsHandledOnlyByTheTreeInterceptor(self) -> None:
		nextHandler = Mock()
		self.interceptor.event_gainFocus(self.window, nextHandler)
		nextHandler.assert_not_called()

	def test_regainingFocusRefreshesSlideBeforeReporting(self) -> None:
		self.interceptor.hadFocusOnce = True
		self.window.handleSlideChange.return_value = True
		with patch.object(powerpnt.SlideShowTreeInterceptor, "reportNewSlide") as reportNewSlide:
			self.interceptor.event_treeInterceptor_gainFocus()
		self.window.handleSlideChange.assert_called_once_with(report=False)
		reportNewSlide.assert_called_once_with(False, resetSelection=True)

	def test_regainingFocusOnSameSlidePreservesReadingPosition(self) -> None:
		self.interceptor.hadFocusOnce = True
		self.window.handleSlideChange.return_value = False
		with patch.object(powerpnt.SlideShowTreeInterceptor, "reportNewSlide") as reportNewSlide:
			self.interceptor.event_treeInterceptor_gainFocus()
		reportNewSlide.assert_called_once_with(True, resetSelection=False)

	def test_newSlideResetsReadingPositionAndUpdatesBrailleWithoutSayAll(self) -> None:
		info = Mock(isCollapsed=True)
		self.interceptor.makeTextInfo = Mock(return_value=info)
		with (
			patch.object(powerpnt.config, "conf", {"virtualBuffers": {"autoSayAllOnPageLoad": False}}),
			patch.object(
				powerpnt.SlideShowTreeInterceptor,
				"selection",
				new_callable=PropertyMock,
			) as selection,
			patch.object(powerpnt.braille.handler, "handleGainFocus") as handleGainFocus,
			patch.object(powerpnt.speech, "speakTextInfo"),
		):
			selection.return_value = info
			self.interceptor.reportNewSlide()
		info.updateCaret.assert_called_once_with()
		handleGainFocus.assert_called_once_with(self.interceptor)


class TestTextFrameFormatting(unittest.TestCase):
	def setUp(self) -> None:
		self.text = "First\rSecond\x0bline\rThird"
		self.paragraphs = [
			SimpleNamespace(
				start=start,
				length=len(text),
				text=text,
				paragraphFormat=SimpleNamespace(
					alignment=alignment,
					bullet=SimpleNamespace(type=powerpnt.ppBulletNumbered, number=number),
				),
			)
			for start, text, alignment, number in [
				(1, "First\r", 1, 1),
				(7, "Second\x0bline\r", 2, 2),
				(19, "Third", 3, 3),
			]
		]
		self.run = SimpleNamespace(start=1, length=len(self.text), font=SimpleNamespace())
		self.textRange = Mock(length=len(self.text))
		self.textRange.runs.return_value = [self.run]
		self.textRange.paragraphs.side_effect = lambda *args: (
			self.paragraphs[args[0] - 1] if args else self.paragraphs
		)
		self.textRange.characters.side_effect = lambda start, length=1: SimpleNamespace(
			text=self.text[start - 1 : start - 1 + length],
			font=self.run.font,
		)
		self.frame = _object(powerpnt.TextFrame, ppObject=SimpleNamespace(textRange=self.textRange))
		self.info = _object(
			powerpnt.TextFrameTextInfo,
			obj=self.frame,
			_startOffset=0,
			_endOffset=len(self.text),
		)
		self.formatConfig = dict.fromkeys(
			[
				"reportFontName",
				"reportFontSize",
				"fontAttributeReporting",
				"reportSuperscriptsAndSubscripts",
				"reportColor",
				"reportLinks",
				"reportAlignment",
			],
			False,
		)

	def test_formattingUsesNumericOffsetAndStopsAtParagraphBoundary(self) -> None:
		field, bounds = self.info._getFormatFieldAndOffsets(6, self.formatConfig)
		self.assertEqual(bounds, (6, 18))
		self.assertEqual(field["line-prefix"], "2.")
		self.assertTrue(field["line-prefix_speakAlways"])

	def test_sameFontRunReportsTheCorrectNumberForEveryParagraph(self) -> None:
		for offset, number in [(0, 1), (6, 2), (18, 3)]:
			with self.subTest(offset=offset):
				field, _bounds = self.info._getFormatFieldAndOffsets(offset, self.formatConfig)
				self.assertEqual(field["line-prefix"], f"{number}.")

	def test_bulletIsNotRepeatedInMiddleOfParagraphOrAfterSoftLineBreak(self) -> None:
		for offset in (2, 13):
			with self.subTest(offset=offset):
				field, _bounds = self.info._getFormatFieldAndOffsets(offset, self.formatConfig)
				self.assertNotIn("line-prefix", field)

	def test_alignmentIsReadFromTheCurrentParagraph(self) -> None:
		self.formatConfig["reportAlignment"] = True
		for offset, alignment in [(0, "left"), (6, "center"), (18, "right")]:
			with self.subTest(offset=offset):
				field, _bounds = self.info._getFormatFieldAndOffsets(offset, self.formatConfig)
				self.assertEqual(field["text-align"], alignment)

	def test_allPowerPointAlignmentsAndUnknownValues(self) -> None:
		self.formatConfig["reportAlignment"] = True
		for value, alignment in [(4, "justify"), (5, "distribute"), (6, "distribute"), (7, "justify")]:
			with self.subTest(value=value):
				self.paragraphs[0].paragraphFormat.alignment = value
				field, _bounds = self.info._getFormatFieldAndOffsets(0, self.formatConfig)
				self.assertEqual(field["text-align"], alignment)
		for value in (-2, 0, 99):
			with self.subTest(value=value):
				self.paragraphs[0].paragraphFormat.alignment = value
				field, _bounds = self.info._getFormatFieldAndOffsets(0, self.formatConfig)
				self.assertNotIn("text-align", field)

	def test_alignmentReportingCanBeDisabled(self) -> None:
		field, _bounds = self.info._getFormatFieldAndOffsets(6, self.formatConfig)
		self.assertNotIn("text-align", field)

	def test_formattingWithoutCalculatingRunsStillUsesCurrentParagraph(self) -> None:
		field, _bounds = self.info._getFormatFieldAndOffsets(18, self.formatConfig, calculateOffsets=False)
		self.assertEqual(field["line-prefix"], "3.")
		self.textRange.runs.assert_not_called()
		self.textRange.characters.assert_called_once_with(19, 1)

	def test_textRangeNormalizesParagraphAndSoftLineBreaks(self) -> None:
		self.assertEqual(self.info._getTextRange(0, len(self.text)), "First\nSecond\nline\nThird")

	def test_invalidRangeRaisesValueError(self) -> None:
		with self.assertRaises(ValueError):
			self.info._getPptTextRange(5, 2)

	def test_linkAtCaretWithoutAFormattingRunReturnsNone(self) -> None:
		self.info.obj.documentWindow = SimpleNamespace(
			ppSelection=SimpleNamespace(textRange=SimpleNamespace(start=len(self.text) + 1)),
		)
		self.assertIsNone(self.info._getLinkDataAtCaretPosition())

	def test_linkDataUsesTheSharedLinkType(self) -> None:
		self.info.obj.documentWindow = SimpleNamespace(
			ppSelection=SimpleNamespace(textRange=SimpleNamespace(start=1)),
		)
		self.run.actionSettings = lambda activation: SimpleNamespace(
			action=powerpnt.ppActionHyperlink,
			Hyperlink=SimpleNamespace(TextToDisplay="NV Access", Address="https://www.nvaccess.org"),
		)
		link = self.info._getLinkDataAtCaretPosition()
		self.assertEqual(link.displayText, "NV Access")
		self.assertEqual(link.destination, "https://www.nvaccess.org")


class TestPowerPointLayoutFeedback(unittest.TestCase):
	def setUp(self) -> None:
		self.shape = _object(
			powerpnt.Shape,
			ppObject=SimpleNamespace(width=72, height=144, rotation=15, lockAspectRatio=-1),
		)

	def test_distancesUseWindowsMeasurementSystem(self) -> None:
		with patch.object(powerpnt.languageHandler, "useImperialMeasurements", return_value=True):
			self.assertEqual(powerpnt.getDistanceText(72), "1 inch")
			self.assertEqual(powerpnt.getDistanceText(144), "2 inches")
		with patch.object(powerpnt.languageHandler, "useImperialMeasurements", return_value=False):
			self.assertEqual(powerpnt.getDistanceText(72), "2.54 centimetres")
			self.assertEqual(powerpnt.getDistanceText(0), "0 centimetres")

	def test_locationQueryIncludesWidthAndHeight(self) -> None:
		with (
			patch.object(powerpnt.languageHandler, "useImperialMeasurements", return_value=True),
			patch.object(powerpnt.Shape, "_getOverlapText", return_value=""),
			patch.object(powerpnt.Shape, "_getShapeLocationText", return_value=""),
		):
			self.assertEqual(self.shape.locationText, "width 1 inch, height 2 inches")

	def test_edgeAndOverlapFeedbackUseTheSameLocalUnits(self) -> None:
		self.shape._edgeDistances = (72, 144, -36, 0)
		self.shape._overlapInfo = (
			None,
			{
				"otherIsBehind": False,
				"label": "rectangle",
				"overlapsOtherLeftBy": 72,
				"overlapsOtherTopBy": 0,
				"overlapsOtherRightBy": 0,
				"overlapsOtherBottomBy": 0,
			},
		)
		with patch.object(powerpnt.languageHandler, "useImperialMeasurements", return_value=True):
			self.assertEqual(self.shape._getOverlapText(), "behind left of rectangle by 1 inch")
			self.assertEqual(
				self.shape._getShapeLocationText(True, True, True, True),
				"Off right slide edge by 0.5 inches, 1 inch from left slide edge, "
				"2 inches from top slide edge, 0 inches from bottom slide edge",
			)

	def test_lockedAspectRatioResizeReportsBothAxes(self) -> None:
		gesture = Mock(modifierNames=["shift"])
		with (
			patch.object(powerpnt.scriptHandler, "isScriptWaiting", return_value=False),
			patch.object(powerpnt.Shape, "_getOverlapText", return_value=""),
			patch.object(powerpnt.Shape, "_getShapeLocationText", return_value="") as location,
		):
			self.shape.script_moveHorizontal(gesture)
			location.assert_called_with(left=True, right=True, top=True, bottom=True)
			self.shape.script_moveVertical(gesture)
			location.assert_called_with(top=True, bottom=True, left=True, right=True)
		self.assertEqual(gesture.send.call_count, 2)

	def test_ordinaryMoveStillReportsOnlyTheMovedAxis(self) -> None:
		gesture = Mock(modifierNames=[])
		with (
			patch.object(powerpnt.scriptHandler, "isScriptWaiting", return_value=False),
			patch.object(powerpnt.Shape, "_getOverlapText", return_value=""),
			patch.object(powerpnt.Shape, "_getShapeLocationText", return_value="") as location,
		):
			self.shape.script_moveHorizontal(gesture)
			location.assert_called_once_with(left=True, right=True, top=False, bottom=False)

	def test_rotationReadsTheAngleAfterSendingTheNativeShortcut(self) -> None:
		gesture = Mock()
		gesture.send.side_effect = lambda: setattr(self.shape.ppObject, "rotation", 30)
		with (
			patch.object(powerpnt.scriptHandler, "isScriptWaiting", return_value=False),
			patch.object(powerpnt.ui, "message") as message,
		):
			self.shape.script_rotate(gesture)
		gesture.send.assert_called_once_with()
		message.assert_called_once_with("rotation 30 degrees")

	def test_hiddenSlideDescriptionDoesNotMakeTheObjectInvisible(self) -> None:
		slide = _object(
			powerpnt.Slide,
			ppObject=SimpleNamespace(slideShowTransition=SimpleNamespace(hidden=-1)),
		)
		self.assertEqual(slide.description, "hidden slide")
		slide.ppObject.slideShowTransition.hidden = 0
		slide.invalidateCache()
		self.assertIsNone(slide.description)

	def test_onlyNonEmptyNotesAddHasNoteState(self) -> None:
		frame = SimpleNamespace(textRange=SimpleNamespace(text="Speaker notes\r"))
		slide = _object(powerpnt.Slide, ppObject=SimpleNamespace())
		with (
			patch.object(
				powerpnt.Window,
				"states",
				new_callable=PropertyMock,
				return_value=set(),
				create=True,
			),
			patch.object(powerpnt, "getNotesTextFrame", return_value=frame),
		):
			self.assertIn(powerpnt.controlTypes.State.HASNOTE, slide._get_states())
			frame.textRange.text = "\r "
			slide.invalidateCache()
			self.assertNotIn(powerpnt.controlTypes.State.HASNOTE, slide._get_states())


class TestSlideShowLocations(unittest.TestCase):
	setUp = TestSlideTitles.setUp

	def test_slideRectangleFitsLetterboxedWindowOnAnotherMonitor(self) -> None:
		self.window.ppObjectModel = SimpleNamespace(
			presentation=SimpleNamespace(pageSetup=SimpleNamespace(slideWidth=960, slideHeight=540)),
		)
		with (
			patch.object(powerpnt.winUser, "getClientRect", return_value=powerpnt.RectLTRB(0, 0, 1200, 900)),
			patch.object(powerpnt.winUser, "ClientToScreen", return_value=(-1200, 100)),
		):
			self.assertEqual(
				self.window.slideRectToScreen(96, 54, 192, 108),
				powerpnt.RectLTRB(-1080, 280, -840, 415),
			)

	def test_invalidViewportDoesNotReturnFabricatedCoordinates(self) -> None:
		self.window.ppObjectModel = SimpleNamespace(
			presentation=SimpleNamespace(pageSetup=SimpleNamespace(slideWidth=960, slideHeight=540)),
		)
		with (
			patch.object(powerpnt.winUser, "getClientRect", return_value=powerpnt.RectLTRB(0, 0, 0, 0)),
			patch.object(powerpnt.winUser, "ClientToScreen", return_value=(0, 0)),
			self.assertRaises(LookupError),
		):
			self.window.slideRectToScreen(0, 0, 10, 10)

	def test_textLocationsAccountForLabelsBulletsAndUtf16Offsets(self) -> None:
		range = Mock(text="A😀B\r", boundLeft=10, boundTop=20, boundWidth=30, boundHeight=40)
		range.characters.return_value = range
		paragraph = range
		paragraph.paragraphFormat.bullet.type = powerpnt.ppBulletNumbered
		paragraph.paragraphFormat.bullet.number = 1
		self.bodyShape.textFrame.textRange.paragraphs = lambda: [paragraph]
		text = self.window.basicText
		info = _object(
			powerpnt.SlideShowTreeInterceptorTextInfo,
			obj=SimpleNamespace(rootNVDAObject=self.window),
		)
		with patch.object(
			powerpnt.SlideShowWindow,
			"slideRectToScreen",
			return_value="screenRect",
		) as convert:
			self.assertEqual(info._getBoundingRectFromOffset(text.index("😀")), "screenRect")
			range.characters.assert_called_with(2, 2)
			info._getBoundingRectFromOffset(text.index("B", text.index("Slide contents")))
			range.characters.assert_called_with(4, 1)
			info._getBoundingRectFromOffset(text.index("1. "))
			range.characters.assert_called_with(1, 1)
			convert.assert_called_with(10, 20, 30, 40)

	def test_graphicDescriptionUsesItsShapeRectangle(self) -> None:
		shape = SimpleNamespace(left=10, top=20, width=100, height=80)
		chunk = powerpnt._SlideTextChunk("Diagram", shape)
		self.window.basicText = "Diagram"
		self.window.basicTextLocations = [(0, 7, chunk)]
		info = _object(
			powerpnt.SlideShowTreeInterceptorTextInfo,
			obj=SimpleNamespace(rootNVDAObject=self.window),
		)
		with patch.object(
			powerpnt.SlideShowWindow,
			"slideRectToScreen",
			return_value="screenRect",
		) as convert:
			self.assertEqual(info._getBoundingRectFromOffset(3), "screenRect")
			convert.assert_called_once_with(10, 20, 100, 80)

	def test_virtualSectionLabelHasNoOnScreenRectangle(self) -> None:
		text = self.window.basicText
		info = _object(
			powerpnt.SlideShowTreeInterceptorTextInfo,
			obj=SimpleNamespace(rootNVDAObject=self.window),
		)
		with self.assertRaises(LookupError):
			info._getBoundingRectFromOffset(text.index("Slide contents"))

	def test_providerFailureBecomesLocationUnavailable(self) -> None:
		range = Mock(text="text")
		range.characters.side_effect = comtypes.COMError(-1, "Unavailable", None)
		self.window.basicText = "text"
		self.window.basicTextLocations = [(0, 4, powerpnt._SlideTextChunk("text", self.bodyShape, range))]
		info = _object(
			powerpnt.SlideShowTreeInterceptorTextInfo,
			obj=SimpleNamespace(rootNVDAObject=self.window),
		)
		with self.assertRaises(LookupError) as caught:
			info._getBoundingRectFromOffset(0)
		self.assertIsInstance(caught.exception.__cause__, comtypes.COMError)


class TestPowerPointTables(unittest.TestCase):
	def setUp(self) -> None:
		self.ppTable = SimpleNamespace(
			firstRow=True,
			firstCol=True,
			rows=SimpleNamespace(count=3),
			columns=SimpleNamespace(count=3),
			cell=Mock(
				return_value=SimpleNamespace(
					shape=SimpleNamespace(textFrame=SimpleNamespace(textRange=Mock(text="Header\r"))),
				),
			),
		)
		self.table = _object(powerpnt.Table, ppObject=SimpleNamespace(table=self.ppTable))
		self.cell = _object(powerpnt.TableCell, table=self.table, rowNumber=2, columnNumber=2)
		self.document = Mock()
		self.frame = _object(powerpnt.TableCellTextFrame, parent=self.cell, documentWindow=self.document)

	def test_rowHeaderUsesTheFirstColumnOfTheCurrentRow(self) -> None:
		self.assertEqual(self.cell.rowHeaderText, "Header")
		self.ppTable.cell.assert_called_once_with(2, 1)

	def test_columnHeaderUsesTheFirstRowOfTheCurrentColumn(self) -> None:
		self.assertEqual(self.cell.columnHeaderText, "Header")
		self.ppTable.cell.assert_called_once_with(1, 2)

	def test_formatFlagsControlHeaderReporting(self) -> None:
		self.ppTable.firstRow = self.ppTable.firstCol = False
		self.assertIsNone(self.cell.rowHeaderText)
		self.assertIsNone(self.cell.columnHeaderText)
		self.ppTable.cell.assert_not_called()

	def test_headerCellRoles(self) -> None:
		self.cell.rowNumber = 1
		self.assertEqual(self.cell.role, powerpnt.controlTypes.Role.TABLECOLUMNHEADER)
		self.cell.rowNumber, self.cell.columnNumber = 2, 1
		self.assertEqual(self.cell.role, powerpnt.controlTypes.Role.TABLEROWHEADER)

	def test_tableNavigationSelectsAnInsertionPointNotTheCellText(self) -> None:
		self.frame._moveTableCell(1, 0)
		self.ppTable.cell.assert_called_once_with(3, 2)
		self.ppTable.cell.return_value.shape.textFrame.textRange.characters.assert_called_once_with(1, 0)
		self.ppTable.cell.return_value.shape.textFrame.textRange.characters.return_value.select.assert_called_once_with()
		self.document.handleSelectionChange.assert_called_once_with()

	def test_navigationAtEdgeDoesNotChangeSelection(self) -> None:
		self.cell.rowNumber = 3
		with patch.object(powerpnt.ui, "message") as message:
			self.frame._moveTableCell(1, 0)
		message.assert_called_once_with("edge of table")
		self.ppTable.cell.assert_not_called()
		self.document.handleSelectionChange.assert_not_called()

	def test_tableNavigationScriptsUseTheExpectedAxis(self) -> None:
		for script, delta in [
			("nextRow", (1, 0)),
			("previousRow", (-1, 0)),
			("nextColumn", (0, 1)),
			("previousColumn", (0, -1)),
		]:
			with self.subTest(script=script), patch.object(self.frame, "_moveTableCell") as move:
				getattr(self.frame, f"script_{script}")(None)
				move.assert_called_once_with(*delta)
