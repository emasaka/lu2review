#!/usr/bin/env perl
use strict;
use warnings;
use utf8;
use Encode;
use OpenOffice::OODoc;
use File::Basename qw(basename);
use File::Spec::Functions qw(catfile);

my $PAGE_WIDTH_CM = 15.1;
my @IMAGE_SUFFIXES = ('.jpg', '.png');

my $imgdir = 'images';
my $filebase;

sub escape_param {
    my ($text) = @_;
    $text =~ s/[\]]/\\$&/g;
    $text;
}

sub width2pageratio {
    my ($width) = @_;
    if ($width =~ /\A([\d.]+)cm/) {
        sprintf '%.2f%%', (($1 / $PAGE_WIDTH_CM) * 100);
    } else {
        '';
    }
}

sub process_draw {
    my ($doc, $elm) = @_;

    my $imgelm = $doc->getNodeByXPath('//draw:image', $elm) or return '';
    my $imgfile = $imgelm->getAttribute('xlink:href');
    return '' if $imgfile =~ /\.wmf\z/;
    my $imgfile_base = basename($imgfile);
    mkdir $imgdir;
    $doc->raw_export($imgfile, catfile($imgdir, "${filebase}-${imgfile_base}"));
    my $imgname = basename($imgfile_base, @IMAGE_SUFFIXES);

    my $width;
    my $imgframe = $doc->getNodeByXPath('..', $imgelm);
    if ($imgframe->hasTag('draw:frame')) {
        my ($w, $h) = $doc->imageSize($imgframe);
        $width = width2pageratio($w);
    }

    my $txt = decode('utf-8', $doc->getText($elm));
    $txt =~ s/\n//g;
    $txt = escape_param($txt);
    $width ? "//indepimage[$imgname][$txt][width=\"$width\"]\n" : "//indepimage[$imgname][$txt]\n";
}

my $subblock_buffer = '';

sub push_subblock {
    $subblock_buffer .= $_[0];
}

sub flush_subblock {
    my $txt = $subblock_buffer;
    $subblock_buffer = '';
    $txt;
}

sub process_footnote {
    my ($doc, $elm) = @_;

    my $name = $elm->getAttribute('text:id');
    my $fn = elm2txt($doc, $doc->getNodeByXPath('//text:note-body', $elm));
    $fn =~ s/\A\n//;
    chomp($fn);
    $fn = escape_param($fn);
    push_subblock "//footnote[${filebase}_${name}][$fn]\n";
    "@<fn>{${filebase}_${name}}";
}

sub process_link {
    my ($doc, $elm) = @_;
    my $href = $elm->getAttribute('xlink:href');
    my $txt = elm2txt($doc, $elm);
    "@<href>{$href,$txt}";
}

sub get_properties {
    my ($doc, $elm) = @_;
    my $sn = $doc->textStyle($elm);
    my $s = $doc->getStyleElement($sn);
    my @p = $doc->styleProperties($s);
    defined($p[0]) ? @p : ();
}

sub process_span {
    my ($doc, $elm) = @_;
    my %p = get_properties($doc, $elm);
    if ($p{'fo:font-weight'} && $p{'fo:font-weight'} eq 'bold') {
        '@<b>{' . elm2txt($doc, $elm) . '}';
    } else {
        elm2txt($doc, $elm);
    }
}

sub elm2txt {
    my ($doc, $elm) = @_;
    my $result = '';

    for my $node ($elm->getChildNodes) {
        if ($node->isElementNode) {
            if ($node->hasTag('text:note')) {
                $result .= process_footnote($doc, $node);
            } elsif ($node->hasTag('draw:frame')) {
                $result .= process_draw($doc, $node);
            } elsif ($node->hasTag('text:a')) {
                $result .= process_link($doc, $node);
            } elsif ($node->hasTag('text:span')) {
                $result .= process_span($doc, $node);
            } else {
                $result .= elm2txt($doc, $node);
            }
        } else {
            $result .= decode('utf-8', $doc->getText($node));
        }
    }
    $result =~ s/\r//g;
    $result;
}

sub list_type {
    my ($doc, $stylename) = @_;
    my $s = $doc->getStyleElement($stylename, namespace => 'text',
                                  type => 'list-style' );
    $doc->getNodeByXPath('//*[@text:level="1"]', $s)->getName;
}

sub list_level {
    my ($doc, $elm) = @_;
    my @list_indents = (0);

    my %p = get_properties($doc, $elm);
    my $m = $p{'fo:margin-left'};
    return 0 unless $m;
    $m =~ s/cm\z//;
    for my $i (0 .. (scalar(@list_indents) - 1)) {
        # FIXIT: what should I do when $m < $list_indents[$i]?
        return $i if ($m <= $list_indents[$i]);
    }
    push @list_indents, $m;
    return scalar(@list_indents) - 1;
}

sub bulletlist2txt {
    my ($doc, $elm) = @_;
    my $result = '';

    for my $item ($doc->getItemElementList($elm)) {
        my $level = list_level($doc, $item) + 1;;
        $result .=  ' ' . ('*' x $level) . ' ' . elm2txt($doc, $item) . "\n";
    }
    $result;
}

sub numberlist2txt {
    my ($doc, $elm) = @_;
    my $result = '';
    my $i = 1;

    for my $item ($doc->getItemElementList($elm)) {
        $result .= ' ' . $i++ . '. ' . elm2txt($doc, $item) . "\n";
    }
    $result;
}

sub table2txt {
    my ($doc, $elm) = @_;

    my $tablename = $elm->getAttribute('table:name');
    my $result = "//table[$tablename][]{\n";

    my ($rows, $columns) = $doc->getTableSize($elm);
    for my $row (0..($rows - 1)) {
        my @line = ();
        for my $col (0..($columns - 1)) {
            my $txt = elm2txt($doc, $doc->getCellParagraph($elm, $row, $col));
            if ($txt) {
                $txt =~ s/\A\./../;
            } else {
                $txt = '.';
            }
            push @line, $txt;
        }
        $result .=  join("\t", @line) . "\n";
        $result .= ('-' x 14) . "\n" if ($row == 0 && $line[0] eq '.');
    }
    $result . "//}\n";
}

sub get_style_name {
    my ($doc, $elm) = @_;
    my $sn = $doc->textStyle($elm);
    my $style = $doc->getStyleElement($sn);
    $sn = $style ? $style->getAttribute('style:parent-style-name') : $sn;
    utf8::is_utf8($sn) ? $sn : decode('utf-8', $sn);
}

my $stat = '';

sub flushstat {
    my ($outfh) = @_;

    if ($stat eq 'code' || $stat eq 'quote' || $stat eq 'author') {
        print $outfh "//}\n\n"
    }
    $stat = '';
}

sub parse_doc {
    my ($doc, $outfh) = @_;

    for my $elm ($doc->getTextElementList) {
        my $sn = get_style_name($doc, $elm);

        if ($elm->hasTag('text:list')) {
            if (list_type($doc, $sn) eq 'text:list-level-style-bullet') {
                print $outfh bulletlist2txt($doc, $elm);
            } else {
                print $outfh numberlist2txt($doc, $elm);
            }
        } elsif ($elm->hasTag('table:table')) {
            print $outfh table2txt($doc, $elm);
        } elsif ($sn eq 'Title') {
            flushstat($outfh);
            print $outfh '= ', elm2txt($doc, $elm), "\n";
        } elsif ($sn eq 'Subtitle') {
            # //subtitle is an original tag
            flushstat($outfh);
            print $outfh "//subtitle{\n", elm2txt($doc, $elm), "\n//}\n";
        } elsif ($sn =~ /\AHeading/) {
            flushstat($outfh);
            my $txt = elm2txt($doc, $elm);
            if ($txt) {
                my $level = ($sn =~ /\AHeading_.*_(\d+)\z/) ? $1 + 1 : 2;
                my $hdr = '=' x $level;
                print $outfh "$hdr $txt\n\n";
            }
        } elsif ($sn eq 'プログラムコード') {
            print $outfh "//emlist{\n" unless $stat eq 'code';
            print $outfh elm2txt($doc, $elm), "\n";
            $stat = 'code';
        } elsif ($sn eq 'Signature' || $sn eq '連絡先') {
            # //author is an original tag
            print $outfh "\n//author{\n" unless $stat eq 'author';
            print $outfh elm2txt($doc, $elm), "\n";
            $stat = 'author';
        } elsif ($sn eq 'Quotations') {
            print $outfh "//quote{\n" unless $stat eq 'quote';
            print $outfh elm2txt($doc, $elm), "\n\n";
            $stat = 'quote';
        } elsif ($sn =~ /\A(?:Text_.*_body|Standard)\z/) {
            flushstat($outfh);
            print $outfh elm2txt($doc, $elm), "\n\n";
        } else {
            # unknown style
            flushstat($outfh);
            print $outfh "<<$sn>>";
            print $outfh elm2txt($doc, $elm), "\n";
        }

        my $block =  flush_subblock();
        print $outfh $block, "\n" if $block;
    }
}

my $odtfile = $ARGV[0] or die;
$filebase = basename($odtfile, '.odt');

open my $outfh, '>:encoding(utf-8)', "${filebase}.re" or die;

my $container = odfContainer($odtfile) or die;
my $doc = odfDocument(container => $odtfile, part => 'content') or die;
parse_doc($doc, $outfh);

close $outfh;
