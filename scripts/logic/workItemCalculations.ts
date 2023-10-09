export class WorkItemCalculations {
    private arrayOfWorkItems = [];

    public constructor(workItems = []) {
        this.arrayOfWorkItems = workItems;
        //console.log('WorkItems array: ' + JSON.stringify(workItems));
    }

    public getWorkItemsResults() {
        var storyPoints = this.getStoryPoints();
        var backendOnlyPoints = this.getTaggedStoryPoints(["Backend"], "Frontend");
        var frontendOnlyPoints = this.getTaggedStoryPoints(["Frontend"], "Backend");
        var combinedPoints = this.getTaggedStoryPoints(["Frontend", "Backend"]);
        var taskNumbers = this.getTaskEstimatedCompletedRemaining();

        var formattedText = this.formatText(storyPoints, taskNumbers, backendOnlyPoints, frontendOnlyPoints, combinedPoints);

        return formattedText;
    }

    private formatText(
        storyPoints: { storiesCount: number; storiesPoints: number; nicsCount: number; nicsPoints: number; featuresCount: number; featuresPoints: number; }, 
        taskNumbers: { tasksCount: number; estimate: number; completed: number; remaining: number; }, 
        backendOnlyPoints: { storiesCount: number; storiesPoints: number; nicsCount: number; nicsPoints: number; featuresCount: number; featuresPoints: number; }, 
        frontendOnlyPoints: { storiesCount: number; storiesPoints: number; nicsCount: number; nicsPoints: number; featuresCount: number; featuresPoints: number; }, 
        combinedPoints: { storiesCount: number; storiesPoints: number; nicsCount: number; nicsPoints: number; featuresCount: number; featuresPoints: number; }) {

        var formattedText =
            (storyPoints.featuresCount > 0 ? storyPoints.featuresCount + " Features: " + storyPoints.featuresPoints + "sp\n" : "") +
            (storyPoints.storiesCount > 0 ? storyPoints.storiesCount + " Stories: " + storyPoints.storiesPoints + "sp\n" : "") +
            (storyPoints.nicsCount > 0 ? storyPoints.nicsCount + " NICs: " + storyPoints.nicsPoints + "sp\nTotal stories + NICs: " + (storyPoints.nicsPoints + storyPoints.storiesPoints) + "sp\n" : "");

        if (frontendOnlyPoints.storiesCount > 0) {
            formattedText += "--------------------------\n" +
            (frontendOnlyPoints.featuresCount > 0 ? frontendOnlyPoints.featuresCount + " FE only Features: " + frontendOnlyPoints.featuresPoints + "sp\n" : "") +
            (frontendOnlyPoints.storiesCount > 0 ? frontendOnlyPoints.storiesCount + " FE only Stories: " + frontendOnlyPoints.storiesPoints + "sp\n" : "") +
            (frontendOnlyPoints.nicsCount > 0 ? frontendOnlyPoints.nicsCount + " FE only NICs: " + frontendOnlyPoints.nicsPoints + "sp\nTotal FE only stories + FE NICs: " + (frontendOnlyPoints.nicsPoints + frontendOnlyPoints.storiesPoints) + "sp\n" : "");
        }

        if (backendOnlyPoints.storiesCount > 0) {
            formattedText += "--------------------------\n" +
            (backendOnlyPoints.featuresCount > 0 ? backendOnlyPoints.featuresCount + " BE only Features: " + backendOnlyPoints.featuresPoints + "sp\n" : "") +
            (backendOnlyPoints.storiesCount > 0 ? backendOnlyPoints.storiesCount + " BE only Stories: " + backendOnlyPoints.storiesPoints + "sp\n" : "") +
            (backendOnlyPoints.nicsCount > 0 ? backendOnlyPoints.nicsCount + " BE only NICs: " + backendOnlyPoints.nicsPoints + "sp\nTotal BE only stories + BE only NICs: " + (backendOnlyPoints.nicsPoints + backendOnlyPoints.storiesPoints) + "sp\n" : "");
        }

        if (combinedPoints.storiesCount > 0) {
            formattedText += "--------------------------\n" +
            (combinedPoints.featuresCount > 0 ? combinedPoints.featuresCount + " Combined Features: " + combinedPoints.featuresPoints + "sp\n" : "") +
            (combinedPoints.storiesCount > 0 ? combinedPoints.storiesCount + " Combined Stories: " + combinedPoints.storiesPoints + "sp\n" : "") +
            (combinedPoints.nicsCount > 0 ? combinedPoints.nicsCount + " Combined NICs: " + combinedPoints.nicsPoints + "sp\nTotal Combined stories + Combined NICs: " + (combinedPoints.nicsPoints + combinedPoints.storiesPoints) + "sp\n" : "");
        }

        if (taskNumbers.tasksCount > 0) {
            formattedText += "--------------------------\n" 
                + taskNumbers.tasksCount + " Tasks:\n"  
                + "Estimated: " + taskNumbers.estimate + "\n"
                + "Completed: " + taskNumbers.completed + "\n"
                + "Remaining: " + taskNumbers.remaining;
        }

        return formattedText.length > 0 ? formattedText : "No PBIs selected";
    }

    private getTaggedStoryPoints(includeTags: string[], excludeTag: string = null) {
        var stories = this.arrayOfWorkItems.filter(
            workitem => includeTags.every(tag => workitem.fields["System.Tags"].includes(tag)) &&
            !workitem.fields["System.Tags"].includes(excludeTag) &&
            (workitem.fields["System.WorkItemType"] == "User Story" ||
            (workitem.fields["System.WorkItemType"] == "Product Backlog Item" && 
            workitem.fields["Roche.DP.VSTS.Complexity"] == "Story")));
        var nics = this.arrayOfWorkItems.filter(workitem => workitem.fields["System.WorkItemType"] == "NIC");
        var features = this.arrayOfWorkItems.filter(workitem => workitem.fields["System.WorkItemType"] == "Feature" ||
            (workitem.fields["System.WorkItemType"] == "Product Backlog Item" && workitem.fields["Roche.DP.VSTS.Complexity"] == "Feature"));

        var storypoints = 0;
        stories.forEach(function (story, index) {
            storypoints = (+story.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + storypoints;
        });

        var nicspoints = 0;
        nics.forEach(function (nic, index) {
            nicspoints = (+nic.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + nicspoints;
        });

        var featpoints = 0;
        features.forEach(function (feature, index) {
            featpoints = (+feature.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + featpoints;
        });

        return {
            storiesCount:  stories.length,
            storiesPoints: storypoints,
            nicsCount:     nics.length,
            nicsPoints: nicspoints,
            featuresCount: features.length,
            featuresPoints: featpoints
        };
    }

    private getStoryPoints() {
        var stories = this.arrayOfWorkItems.filter(
            workitem => workitem.fields["System.WorkItemType"] == "User Story" ||
            (workitem.fields["System.WorkItemType"] == "Product Backlog Item" && workitem.fields["Roche.DP.VSTS.Complexity"] == "Story"));
        var nics = this.arrayOfWorkItems.filter(workitem => workitem.fields["System.WorkItemType"] == "NIC");
        var features = this.arrayOfWorkItems.filter(workitem => workitem.fields["System.WorkItemType"] == "Feature" ||
            (workitem.fields["System.WorkItemType"] == "Product Backlog Item" && workitem.fields["Roche.DP.VSTS.Complexity"] == "Feature"));

        var storypoints = 0;
        stories.forEach(function (story, index) {
            storypoints = (+story.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + storypoints;
        });

        var nicspoints = 0;
        nics.forEach(function (nic, index) {
            nicspoints = (+nic.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + nicspoints;
        });

        var featpoints = 0;
        features.forEach(function (feature, index) {
            featpoints = (+feature.fields["Microsoft.VSTS.Scheduling.StoryPoints"] || 0) + featpoints;
        });

        return {
            storiesCount:  stories.length,
            storiesPoints: storypoints,
            nicsCount:     nics.length,
            nicsPoints: nicspoints,
            featuresCount: features.length,
            featuresPoints: featpoints
        };
    }

    private getTaskEstimatedCompletedRemaining() {
        var tasks = this.arrayOfWorkItems.filter(workitem => workitem.fields["System.WorkItemType"] == "Task");

        var remainingwork = 0, originalestimate = 0, completedwork = 0;
        tasks.forEach(function (task, index) {
            remainingwork = (+task.fields["Microsoft.VSTS.Scheduling.RemainingWork"] || 0) + remainingwork;
            originalestimate = (+task.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0) + originalestimate;
            completedwork = (+task.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0) + completedwork;
        });

        return {
            tasksCount: tasks.length,
            estimate: originalestimate,
            completed: completedwork,
            remaining: remainingwork
        };
    }

}
